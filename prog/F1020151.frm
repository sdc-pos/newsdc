VERSION 5.00
Begin VB.Form F1020151 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出荷予定データ取込み (F102015 2016.03.08 09：30) "
   ClientHeight    =   4170
   ClientLeft      =   1905
   ClientTop       =   2385
   ClientWidth     =   8580
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
   ScaleHeight     =   4170
   ScaleWidth      =   8580
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox LBox_Hin 
      Height          =   300
      Left            =   1560
      TabIndex        =   25
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   6360
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
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
      Index           =   9
      Left            =   5760
      TabIndex        =   19
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
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
      Index           =   8
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
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
      Index           =   7
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
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
      Index           =   6
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
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
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   3
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
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
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
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
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
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
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "F1020151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String * 3           'ﾜｰｸｽﾃｰｼｮﾝ番号

Private FileName    As String               'テキストファイル名
Private FileNo      As Integer              'ファイル№

Private KASO_NYUKA_SOKO      As String * 2  '仮想入荷倉庫番号
Private KASO_SMODOSHI_SOKO   As String * 2  '仮想支給戻し倉庫番号

Private Proc_F      As Integer              '品番＆在庫有無　判定フラグ
Private Last_Proc_F As Integer              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグ
                                            
Private Type YUKO_SOKO_TBL                  '有効ﾎｽﾄ倉庫取り込みテーブル
    HS_SOKO             As String * 8
    NAIGAI              As String * 1
End Type

Dim Soko_T()            As YUKO_SOKO_TBL  '倉庫情報

'-                                          2005.12.30
Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '仕向け先
    JGYOBU                  As String * 1   '事業部
    NAIGAI                  As String * 1   '国内外
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean

'-                                          2005.12.30


Private New_HS_IN_SIJ   As String           '入庫データファイル名
Private New_HS_OUT_SIJ  As String           '新ﾚｲｱｳﾄ出庫データファイル名


Private In_Cnt      As Integer              'データ読み込み件数
Private Out_Cnt     As Integer              'データ出力件数

Private Const In_Mode% = 1                  '入荷処理
Private Const Out_Mode% = 2                 '出荷処理

                                            
Dim NormalFont As New StdFont               '印刷フォント

Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 1                    '明細印刷開始桁位置（１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）
Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private Const NAI_CHANGE% = 1
Private Const GAI_CHANGE% = 2
Private Const NOT_GAI_CHANGE% = 3


Private ETC_MTS_NAI As String * 8             'その他向け先(国内)
Private ETC_SS_NAI  As String * 8             'その他直送先(国内)

Private ETC_MTS_GAI As String * 8             'その他向け先(海外)
Private ETC_SS_GAI  As String * 8             'その他直送先(海外)

Dim DUP_SYUKA_DATA  As String                 '出荷データフルパス

                                        
Dim MyCenter        As String

Dim Err_FLg         As Boolean

Dim TANA_SPACE      As Boolean          '2009.03.07
                                        
Private MENU_NO     As String * 2       '実績ログ出力用ﾒﾆｭｰ№   2007.11.06
                                        
                                        
                                        
Dim RYOHEN_TANA     As String * 8       '良品返品入庫棚番       2011.01.18
                                        
'商品化計画支援 2011.07.07
Dim NOT_Hin_Name    As Variant          '除外品名
Dim NOT_Hin_Name_F  As Boolean          '除外品名有無
'商品化計画支援 2011.07.07
Dim GOODS_F         As String * 1       '商品化有無　ﾃﾞﾌｫﾙﾄ 2012.12.20
                                        
                                        
Private Function New_Nyuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理
'----------------------------------------------------------------------------

''''''''''''''''''''''''''''''' 全て廃止 2009.06.18
'''Dim i           As Integer
'''Dim j           As Integer
'''Dim Skip_Flg    As Boolean
    
'''Dim WK_Y_QTY    As Long     '出荷数ワーク
'''Dim WK_Qty      As Long     '前借残ワーク
'''Dim WK_E_QTY    As Long     '先行出荷数ワーク

'''Dim SUMI_QTY    As Long     '商品化済みとして登録
'''Dim MI_QTY      As Long     '未商品として登録

'''Dim Work_SOKO     As String * 2
    
'''Dim sts         As Integer
'''Dim ans         As Integer
'''Dim Not_SHUSI   As Boolean
    
''''出荷予定 編集前処理 ################################################################# 2005/05/16 Add ↓
'''Dim Fast_Flg        As Boolean
'''Dim DUP_SYUKANo     As Integer
'''Dim fileName        As String
'''Dim Ret             As Integer
'''Dim INS_NOW         As String * 14
'''Dim wkStr           As String
    
'''Dim wkMUKE_CODE     As String
    
    
'''Dim NAIGAI          As String * 1
    
'''''''''''''''''''''''''''''   全ての入荷処理を廃止    2007.06.22
'''    Fast_Flg = True
'''
'''
'''
'''    DUP_SYUKANo = FreeFile
'''    fileName = DUP_SYUKA_DATA
'''
'''    Ret = InStr(1, Trim(fileName), ".") - 1
'''    fileName = Left(Trim(fileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
'''
'''    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
''''#################################################################################### 2005/05/16 Add ↑
'''
'''    New_Nyuka_Update_Proc = True
'''
'''
'''    Do
'''        Get #FileNo, , New_HS_IN_SIJREC
'''        If Left(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), 1) < " " Then
'''            Exit Do
'''        End If
'''
'''        If StrConv(New_HS_IN_SIJREC.CR_LF, vbUnicode) <> vbCrLf Then
'''
'''            Call NG_File_Make_Proc
'''
'''
'''            Exit Do
'''        End If
'''
'''
'''        In_Cnt = In_Cnt + 1
'''        lblINCNT(i).Caption = Format(In_Cnt, "#0")
'''        DoEvents
'''
'''
''''        Skip_Flg = True
''''        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
''''            If JGYOBU = JGYOBU_T(i).CODE Then
''''                For j = 0 To UBound(Soko_T, 2)
''''                    If StrConv(HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
''''                        Skip_Flg = False
''''                        Exit For
''''                    End If
''''                Next j
''''                Exit For
''''            End If
''''        Next i
'''
''''--     2005.12.30
'''        Skip_Flg = True
'''        Not_SHUSI = False
'''        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
'''            If JGYOBU = JGYOBU_T(i).CODE Then
'''                For j = 0 To UBound(Soko_T, 2)
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
'''                        Skip_Flg = False
'''                        Exit For
'''                    End If
'''                Next j
'''                Exit For
'''            End If
'''        Next i
'''
'''        If Skip_Flg Then
'''            Not_SHUSI = True
'''        End If
''''--     2005.12.30
'''
'''
'''
'''
'''
'''        If StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode) <> "1" Then
'''            Skip_Flg = True
'''        End If
'''
'''
'''        If StrConv(New_HS_IN_SIJREC.PM_KBN, vbUnicode) = "-" Then
'''            Skip_Flg = True
'''        End If
'''
'''        'NOPOS  2006.05.01
'''        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "NOPOS" Then
'''            Skip_Flg = True
'''        End If
'''
'''
'''
'''
'''
'''
'''        Work_SOKO = KASO_NYUKA_SOKO
'''
'''
'''
'''        Select Case JGYOBU
'''
'''''            Case SENTAKU                        '洗濯機
'''''
'''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'''''                    Skip_Flg = True
'''''                End If
'''''
'''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'''''                    Skip_Flg = True
'''''                End If
'''
'''
'''
'''
'''            Case SOJIKI                         '掃除機
'''
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KM" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "GG" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "SS" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                '2005.04.07 収支追加
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 5) = "0090K" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KM" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KM" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "GG" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "GG" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "SS" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "SS" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'''                    Work_SOKO = KASO_SMODOSHI_SOKO
'''                End If
'''
'''
'''
'''            Case DENKA, SUIHAN, SENTAKU         '電化、炊飯、洗濯機（アイロン）
'''
'''
'''                Select Case MyCenter
'''
'''                    Case "O"
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "01" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "H33" Then    '2004.07.16
'''                            Skip_Flg = True
'''                        End If
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "H22" Then    '2004.07.16
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "05" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "G22" Then
'''                            Work_SOKO = "80"
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "G11" Then
'''                            Work_SOKO = "81"
'''                        End If
'''
'''                        '2006.04.29用
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "S1" And _
'''                            Left(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "S3" Then
'''                            Work_SOKO = "87"
'''                        End If
'''                        '2006.05.01
'''                        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "POS87" Then
'''                            Work_SOKO = "87"
'''                        End If
'''
'''
'''                    Case "F"
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) <> "904" Then
'''                            If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'''                              Skip_Flg = True
'''                            End If
'''                        End If
'''
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "S1" And _
'''                            Left(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "S2" Then
'''                            Work_SOKO = "88"
'''                        End If
'''
'''                        '2006.05.01
'''                        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "POS88" Then
'''                            Work_SOKO = "88"
'''                        End If
'''
'''
'''                End Select
'''             Case AIRCON                     'エアコン
'''
'''                If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "J4" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JG" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JW" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JV" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "HY" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) = "SH" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) = "S1" Then
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "OS" Then
'''                      Skip_Flg = True
'''                    End If
'''                End If
'''
'''                If Not Skip_Flg Then
''''---------------    2005.06.14
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
'''                        Work_SOKO = "80"
'''                    Else
'''                        If StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode) = "A" Then
'''                            If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                                Work_SOKO = "92"
'''                            Else
'''                                Work_SOKO = "95"
'''                            End If
'''                        Else
'''                            If StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode) = "D" Then
'''                                Work_SOKO = "70"
'''                            Else
''''---------------    2005.06.14
'''                                If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                                    Work_SOKO = "92"
'''                                Else
'''
''''                                    If StrConv(HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
''''                                        Work_SOKO = "80"
''''                                    Else
'''                                        Work_SOKO = "95"
''''                                    End If
'''                                End If
'''                            End If
'''                        End If
'''                    End If
'''                End If
'''
'''
'''
'''        End Select
'''
'''
'''
'''
'''
'''
'''
'''        If Not Skip_Flg Then
'''
'''
'''
'''
'''                                        '入荷予定重複チェック
'''            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'''            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''            Select Case sts
'''                Case BtNoErr
'''                    Call Log_Out(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''                    Skip_Flg = True
'''                Case BtErrKeyNotFound
'''                Case Else
'''                    Call File_Error(sts, BtOpGetEqual, "入荷予定")
'''                    Exit Function
'''            End Select
'''
'''            If Not Skip_Flg Then
'''                                                'トランザクション開始
'''                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                If sts <> BtNoErr Then
'''                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
'''                    Exit Function
'''                End If
'''                                            '品目マスタチェック
'''                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode)) Then
'''                    GoTo Abort_Tran
'''                End If
'''
'''
'''                                            '入荷データ作成
'''                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
'''                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'''                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'''                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
'''                Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''
'''
'''
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
'''                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.KAIKEI_JGYOBA, "")
'''                Call UniCode_Conv(Y_NYUREC.SHISAN_JGYOBA, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'''                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.HOJYO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.TANKA, "")
'''                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
'''                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
'''                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
'''                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
'''                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
'''                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
'''                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
'''                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
'''                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.S_SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.S_HOJYO_SYUSI, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
'''                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAN_SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAN_HOJYO_SYUSI, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
'''                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.SS_CODE, "")
'''                Call UniCode_Conv(Y_NYUREC.KEPIN_KAIJYO, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'''
'''
'''                Last_Proc_F = True              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り
'''
'''
'''                '入荷ﾁｪｯｸﾃﾞｰﾀ更新
'''                Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
'''                Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
'''                Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''
'''                WK_Y_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''
'''
'''                Do
'''                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                    Select Case sts
'''                        Case BtNoErr
'''                            If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
'''                                WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
'''                                Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
'''
'''                                Do
'''
'''                                    sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ")
'''                                            Exit Function
'''                                    End Select
'''
'''                                Loop
'''                                WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            Else
'''                                Do
'''                                    sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''                                WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
'''                            End If
'''
'''                            Exit Do
'''                        Case BtErrKeyNotFound
'''                            WK_E_QTY = 0
'''                            Exit Do
'''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                            Beep
'''                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                            If ans = vbCancel Then
'''                                Exit Function
'''                            End If
'''                        Case Else
'''                            Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ")
'''                            Exit Function
'''                    End Select
'''                Loop
'''                                    '先行入荷数（入荷実績数）
'''                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
'''
'''                                    '予算単位元
'''                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'''                                    '予算単位先
'''                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''                                    '標準棚番
'''                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''
'''                Call UniCode_Conv(Y_NYUREC.FILLER, "")
'''
'''                Do
'''                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                    Select Case sts
'''                        Case BtNoErr
'''                            Exit Do
'''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                            Beep
'''                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                            If ans = vbCancel Then
'''                                Exit Function
'''                            End If
'''                        Case Else
'''                            Call File_Error(sts, BtOpInsert, "入荷予定")
'''                            Exit Function
'''                    End Select
'''                Loop
'''
''''------------ 2005.12.30
'''                Select Case JGYOBU
'''                    Case AIRCON, SENTAKU
'''                        Call UniCode_Conv(K0_SOKO.Soko_No, Work_SOKO)
'''                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'''                        Select Case sts
'''                            Case BtNoErr
'''                            Case Else
'''                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'''                                Exit Function
'''                        End Select
'''
'''                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
'''
'''                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            MI_QTY = 0
'''                        Else
'''
'''                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                                SUMI_QTY = 0
'''                            Else
'''                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                                MI_QTY = 0
'''                            End If
'''                        End If
'''
''''------------ 2005.12.30
'''
'''                    Case Else
'''
'''
'''                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            SUMI_QTY = 0
'''                        Else
'''                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            MI_QTY = 0
'''                        End If
'''                End Select
'''
'''
''''                Wk_SOKO = KASO_NYUKA_SOKO
''''                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
''''                    Wk_SOKO = KASO_SMODOSHI_SOKO
''''
''''                End If
'''
'''                '入荷数で在庫データ更新（＋）
'''                If Nyuko_Update_Proc(JGYOBU, _
'''                                    Soko_T(i, j).NAIGAI, _
'''                                    StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
'''                                    (Work_SOKO & "01" & "01" & "01"), _
'''                                    YOIN_TU_NYUKA, _
'''                                    SUMI_QTY, MI_QTY, _
'''                                    WS_NO, WS_NO, , _
'''                                    StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode) & " 伝№:" & StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) Then
'''                    Exit Function
'''
'''                End If
'''
'''                '前借り数で在庫データ更新（－）
'''                If WK_E_QTY <> 0 Then
'''                '在庫データLOCK
'''                    If Zaiko_Lock_Proc((Work_SOKO & "01" & "01" & "01"), _
'''                                        JGYOBU, _
'''                                        Soko_T(i, j).NAIGAI, _
'''                                        StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                        WS_NO) Then
'''                        Exit Function
'''
'''                    End If
'''
'''                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                        MI_QTY = WK_E_QTY
'''                    Else
'''                        SUMI_QTY = WK_E_QTY
'''                    End If
'''
'''
'''                    If Syuko_Update_Proc(JGYOBU, _
'''                                        Soko_T(i, j).NAIGAI, _
'''                                        StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                        StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode), _
'''                                        (Work_SOKO & "01" & "01" & "01"), _
'''                                        YOIN_MAE_SOUSAI, _
'''                                        SUMI_QTY, MI_QTY, 0, _
'''                                        WS_NO, WS_NO) Then
'''                        Exit Function
'''
'''                    End If
'''
'''
'''
'''
'''
'''
'''                End If
'''
'''
'''                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                If sts <> BtNoErr Then
'''                    GoTo Abort_Tran
'''                End If
'''
'''
'''                Out_Cnt = Out_Cnt + 1
'''                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
'''                DoEvents
'''
'''            End If
'''
'''
'''
'''
'''
''''出荷予定変換################################################## 2005/05/16 Add 滋賀物流↓
'''        Else
'''
'''            If JGYOBU = AIRCON Then
'''
'''                If Not_SHUSI Then
'''                Else
'''                    If StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode) = "2" Then
'''
'''
'''                        wkMUKE_CODE = ""
'''
'''                        If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
'''                            wkMUKE_CODE = "S8"
'''                        Else
'''                            If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                            Else
'''                                Select Case Trim(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''
'''                                    Case "Z0014"
'''                                        wkMUKE_CODE = "LM"
'''                                    Case "B0070"
'''                                        If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S2" Then
'''                                            wkMUKE_CODE = "S2"
'''                                        Else
'''                                            wkMUKE_CODE = "AC"
'''                                        End If
'''                                    Case Else
'''                                        wkMUKE_CODE = "AC"
'''                                End Select
'''                            End If
'''                        End If
'''
'''
'''                        If wkMUKE_CODE = "" Then
'''                        Else
'''                            Skip_Flg = False
'''                                                        '入荷予定重複チェック
'''                            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'''                            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''                            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                            Select Case sts
'''                                Case BtNoErr
'''                                    Call Log_Out(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''                                    Skip_Flg = True
'''                                Case BtErrKeyNotFound
'''                                Case Else
'''                                    Call File_Error(sts, BtOpGetEqual, "入荷予定")
'''                                    Exit Function
'''                            End Select
'''
'''
'''
'''
'''
'''        ''                                    '出荷予定重複チェック
'''        ''                    Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
'''        ''                    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'''        ''
'''        ''                    sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
'''        ''                    Select Case sts
'''        ''                        Case BtNoErr
'''        ''                            Call Log_Out(LOG_F, "Y_SYUKA.DAT DUP 事業部=" & JGYOBU & "伝票ＩＤ＝" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'''        ''                            Skip_Flg = True
'''        ''
'''        ''                            If Fast_Flg Then
'''        ''                                Open (fileName) For Output As DUP_SYUKANo
'''        ''                                Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")
'''        ''                                Write #DUP_SYUKANo, "出荷日", "伝票№", "支払先ｺｰﾄﾞ", "倉庫/ＳＳｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"
'''        ''                                Fast_Flg = False
'''        ''                            End If
'''        ''
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.SYUKA_YMD, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.DEN_NO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_NAME, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.CHU_KBN, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.CHU_KBN_NAME, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.SURYO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode)
'''        ''
'''        ''                        Case BtErrKeyNotFound
'''        ''                        Case Else
'''        ''                            Call File_Error(sts, BtOpGetEqual, "出荷予定")
'''        ''                            Exit Function
'''        ''                    End Select
'''
'''                            If Not Skip_Flg Then
'''
'''                                                                'トランザクション開始
'''                                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                                If sts <> BtNoErr Then
'''                                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
'''                                    Exit Function
'''                                End If
'''                                                                '品目マスタチェック
'''                                If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode)) Then
'''                                    GoTo Abort_Tran
'''                                End If
'''
'''        '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
'''                                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
'''                                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'''                                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'''                                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
'''                                Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''
'''
'''
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
'''                                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.KAIKEI_JGYOBA, "")
'''                                Call UniCode_Conv(Y_NYUREC.SHISAN_JGYOBA, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'''                                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.TANKA, "")
'''                                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
'''                                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
'''                                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
'''                                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
'''                                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
'''                                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
'''                                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.S_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.S_HOJYO_SYUSI, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
'''                                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAN_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAN_HOJYO_SYUSI, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
'''                                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.SS_CODE, "")
'''                                Call UniCode_Conv(Y_NYUREC.KEPIN_KAIJYO, "")
'''                                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'''                                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
'''                                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_NYUREC.FILLER, "")
'''
'''                                Do
'''                                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpInsert, "入荷予定")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''        '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
'''
'''                                Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
'''                                Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
'''                                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
'''                                Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
'''                                Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
'''
'''                                If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
'''                                    GoTo Abort_Tran
'''                                Else
'''                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
'''                                End If
'''
'''                                Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
'''                                Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
'''                                Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                                wkStr = Format(Val(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000")
'''                                Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
'''                                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
'''                                Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.TANKA, "")
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
'''                                Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
'''
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_IN_SIJREC.SYUK_NAME, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
'''                                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
'''                                Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
'''                                Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
'''                                Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
'''                                Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_IN_SIJREC.CYOK_KBN, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
'''                                Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
'''                                Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
'''                                Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(New_HS_IN_SIJREC.HOST_TANA, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
'''                                Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
'''
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.FILLER, "")
'''
'''                                Do
'''                                    sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpInsert, "出荷予定")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''
'''                                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                                If sts <> BtNoErr Then
'''                                    GoTo Abort_Tran
'''                                End If
'''
'''                                Out_Cnt = Out_Cnt + 1
'''                                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
'''                                DoEvents
'''
'''                                If SYUKA_LOG_ON Then
'''                                    Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
'''                                End If
'''
'''                                If Not Fast_Flg Then
'''                                    Close #DUP_SYUKANo
'''                                End If
'''                            End If
'''                        End If
'''                    End If
'''                End If
'''            End If
''''#################################################################################### 2005/05/16 Add ↑
'''
'''
'''
'''
'''
'''        End If
'''
'''    Loop
'''    New_Nyuka_Update_Proc = False
'''
'''    Exit Function
'''
'''Abort_Tran:
'''
'''    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''    If sts <> BtNoErr Then
'''        Call File_Error(sts, BtOpAbortTransaction, "")
'''    End If


    
''''''''''''''''''''''''''''''' 全て廃止    2009.06.18
'''    New_Nyuka_Update_Proc = True
'''
'''    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'''
'''
'''
'''
'''    Do
'''        Get #FileNo, , New_HS_IN_SIJREC
'''        If Left(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), 1) < " " Then
'''            Exit Do
'''        End If
'''
'''        If StrConv(New_HS_IN_SIJREC.CR_LF, vbUnicode) <> vbCrLf Then
'''
'''            Call NG_File_Make_Proc
'''
'''
'''            Exit Do
'''        End If
'''
'''
'''        In_Cnt = In_Cnt + 1
'''        lblINCNT(i).Caption = Format(In_Cnt, "#0")
'''        DoEvents
'''
'''
'''        Skip_Flg = True
'''        Not_SHUSI = False
'''        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
'''            If StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
'''                For j = 0 To UBound(Soko_T, 2)
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
'''                        Skip_Flg = False
'''                        Exit For
'''                    End If
'''                Next j
'''                Exit For
'''            End If
'''        Next i
'''
'''        If Skip_Flg Then
'''            Not_SHUSI = True
'''        End If

''''-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
'''        '照合用入荷予定重複チェック
'''        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode))
'''        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'''        Select Case sts
'''            Case BtNoErr
'''                Call Log_Out(LOG_F, "Y_GLICS.DAT DUP 事業部=" & StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode) & "ＴＥＸＴＩＤ＝" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''            Case BtErrKeyNotFound
'''            Case Else
'''                Call File_Error(sts, BtOpGetEqual, "照合用入荷予定")
'''                Exit Function
'''        End Select
'''
'''
'''        If Not_SHUSI Then
'''            NAIGAI = "1"
'''        Else
'''            NAIGAI = Soko_T(i, j).NAIGAI
'''        End If
'''
'''        If sts = BtErrKeyNotFound Then
'''
'''            If Y_GLICS_PUT_PROC(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), NAIGAI, INS_NOW) Then
'''                Exit Function
'''            End If
'''
'''        End If
'''
'''
'''    Loop
''''-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
'''
'''
'''
'''
'''    New_Nyuka_Update_Proc = False
    





'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理  F102010より移行
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Skip_Flg    As Boolean
    
Dim WK_Y_QTY    As Long     '出荷数ワーク
Dim WK_Qty      As Long     '前借残ワーク
Dim WK_E_QTY    As Long     '先行出荷数ワーク

Dim SUMI_QTY    As Long     '商品化済みとして登録
Dim MI_QTY      As Long     '未商品として登録

Dim WORK_SOKO   As String * 2
    
Dim sts         As Integer
Dim ans         As Integer
Dim Not_SHUSI   As Boolean
    
Dim wkText      As String
Dim Length      As Integer
    
    
Dim NAIGAI      As String * 1   '2007.06.15
    
    
Dim TEXT_NO     As String * 9           'ﾃｷｽﾄ№
Dim JGYOBU_Code As String * 1           '事業部区分
Dim CYOK_KBN    As String * 1           '直送区分
Dim DEN_DT      As String * 8           '伝票日付
Dim IO_KBN      As String * 1           '入出庫区分
Dim PM_KBN      As String * 1           '赤黒区分
Dim DEN_SYU     As String * 1           '伝票種別
Dim DEN_NO      As String * 6           '伝票№
Dim CYU_KBN     As String * 1           '注文区分
'Dim HIN_GAI     As String * 13          '品番（外部）  '13-->20 2016.03.07
Dim HIN_GAI13     As String * 13          '品番（外部）   '13-->20 2016.03.07
Dim HIN_GAI20     As String * 20          '品番（外部）   '13-->20 2016.03.07
Dim HIN_GAI     As String * 20          '品番（外部）   '13-->20 2016.03.07


Dim HIN_NAI     As String * 13          '品番（内部）
Dim HIN_NAME    As String * 25          '品名
Dim YOTEI_QTY   As String * 6           '数量
Dim YOSAN_FROM  As String * 5           '予算単位（元）
Dim YOSAN_TO    As String * 5           '予算単位（先）
Dim HOST_SOKO   As String * 8           '倉庫区分（ﾎｽﾄ）
Dim HOST_TANA   As String * 8           '棚番（ﾎｽﾄ）
Dim SYUK_CODE   As String * 5           '支給先／出荷先
Dim SYUK_NAME   As String * 20          '支給先／出荷先名
Dim REC_END     As String * 1           'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    
    
    
    
'2011.01.18
Dim GENSANKOKU          As String * 20  '原産国名
Dim GEN_GENSANKOKU      As String * 20  '現物表示原産国名
Dim SHIIRE_WORK_CENTER  As String * 8   '資材仕入先ﾜｰｸｾﾝﾀｰ
Dim KANKYO_KBN          As String * 3   '環境種類区分
Dim KANKYO_KBN_ST       As String * 8   '環境種類区分適用開始
Dim KANKYO_KBN_SURYO    As String * 10  '環境種類区分数量
Dim ID_NO2              As String * 12  'ID_NO
Dim AITESAKI_CODE       As String * 16  '相手先ｺｰﾄﾞ
Dim JYUCHU_YMD          As String * 8   '受注年月日
Dim SHITEI_NOUKI_YMD    As String * 8   '指定納期年月日


Dim GENSAN_CNT          As Integer

Dim com                 As Integer
'2011.01.18
    
    
    
    
'出荷予定 編集前処理 ################################################################# 2005/05/16 Add ↓
Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
Dim INS_NOW         As String * 14
Dim wkStr           As String
    
Dim wkMUKE_CODE     As String
    
    
Dim Loop_Cnt        As Integer          '2011.01.15

'2011.03.23
Dim MOTO_TEXT_NO    As String * 9
'2011.03.23
Dim DUP_FLG         As Boolean

    
    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'#################################################################################### 2005/05/16 Add ↑
    
    New_Nyuka_Update_Proc = True


    Do Until EOF(FileNo)
        Line Input #FileNo, wkText
    
    
    
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 Then    '138-->145 2016.03.07
        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And _
            LenB(StrConv(wkText, vbFromUnicode)) <> 145 Then     '138-->145 2016.03.07
            
            
            
'            Call NG_File_Make_Proc
             Err_FLg = True
           
            
            Exit Do
        End If
    
    
    
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
    
    
    
    
                                                                    'ﾃｷｽﾄ№
        Length = 1
        TEXT_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TEXT_NO)), vbUnicode)
                                                                    '事業部区分
        Length = Length + Len(TEXT_NO)
        JGYOBU_Code = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JGYOBU_Code)), vbUnicode)
                                                                    '直送区分
        Length = Length + Len(JGYOBU_Code)
        CYOK_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYOK_KBN)), vbUnicode)
                                                                    '伝票日付
        Length = Length + Len(CYOK_KBN)
        DEN_DT = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_DT)), vbUnicode)
                                                                    '入出庫区分
        Length = Length + Len(DEN_DT)
        IO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(IO_KBN)), vbUnicode)
                                                                    '赤黒区分
        Length = Length + Len(IO_KBN)
        PM_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(PM_KBN)), vbUnicode)
                                                                    '伝票種別
        Length = Length + Len(PM_KBN)
        DEN_SYU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_SYU)), vbUnicode)
                                                                    '伝票№
        Length = Length + Len(DEN_SYU)
        DEN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_NO)), vbUnicode)
                                                                    '注文区分
        Length = Length + Len(DEN_NO)
        CYU_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYU_KBN)), vbUnicode)
                                                                    '品番（外部）
        Length = Length + Len(CYU_KBN)
'        HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI)), vbUnicode)
                                                                    '品番（内部）
        If LenB(StrConv(wkText, vbFromUnicode)) = 138 Then
            HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI13)), vbUnicode)
            Length = Length + Len(HIN_GAI13)
        Else
            HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI20)), vbUnicode)
            Length = Length + Len(HIN_GAI20)
        End If
        
        
        HIN_NAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAI)), vbUnicode)
                                                                    '品名
        Length = Length + Len(HIN_NAI)
        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)
                                                                    '数量
        Length = Length + Len(HIN_NAME)
        YOTEI_QTY = Trim(StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOTEI_QTY)), vbUnicode))
                                                                    '予算単位（元）
        Length = Length + Len(YOTEI_QTY)
        YOSAN_FROM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_FROM)), vbUnicode)
                                                                    '予算単位（先）
        Length = Length + Len(YOSAN_FROM)
        YOSAN_TO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_TO)), vbUnicode)
                                                                    '倉庫区分（ﾎｽﾄ）
        Length = Length + Len(YOSAN_TO)
        HOST_SOKO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_SOKO)), vbUnicode)
                                                                    '棚番（ﾎｽﾄ）
        Length = Length + Len(HOST_SOKO)
        HOST_TANA = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_TANA)), vbUnicode)
                                                                    '支給先／出荷先
        Length = Length + Len(HOST_TANA)
        SYUK_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_CODE)), vbUnicode)
                                                                    '支給先／出荷先名
        Length = Length + Len(SYUK_CODE)
        SYUK_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_NAME)), vbUnicode)
    
    
    
        '2011.01.18
        GENSANKOKU = ""             '原産国名
        GEN_GENSANKOKU = ""         '現物表示原産国名
        SHIIRE_WORK_CENTER = ""     '資材仕入先ﾜｰｸｾﾝﾀｰ
        KANKYO_KBN = ""             '環境種類区分
        KANKYO_KBN_ST = ""          '環境種類区分適用開始
        KANKYO_KBN_SURYO = ""       '環境種類区分数量
        ID_NO2 = TEXT_NO            'ID_NO
        AITESAKI_CODE = ""          '相手先ｺｰﾄﾞ
        JYUCHU_YMD = ""             '受注年月日
        SHITEI_NOUKI_YMD = ""       '指定納期年月日
        '2011.01.18
    
        Skip_Flg = True
        Not_SHUSI = False
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If JGYOBU_Code = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(HOST_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        If Skip_Flg Then
            Not_SHUSI = True
        End If
    
    
'-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
        '照合用入荷予定重複チェック
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU_Code)
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)

        DUP_FLG = False                 '2011.03.23

        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP 事業部=" & JGYOBU_Code & "ＴＥＸＴＩＤ＝" & TEXT_NO)
                DUP_FLG = True          '2011.03.23
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "照合用入荷予定", 0)
                Exit Function
        End Select


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 発生元によるデータ確認  2011.03.23
        MOTO_TEXT_NO = ""
        
        If DUP_FLG Then
            If StrConv(App.EXEName, vbUpperCase) = Trim(StrConv(Y_GLICSREC.MOTO_PROG_ID, vbUnicode)) Then
            Else
                If Trim(TEXT_NO) = Trim(StrConv(Y_GLICSREC.MOTO_TEXT_NO, vbUnicode)) Then
                Else
                    MOTO_TEXT_NO = TEXT_NO
                    Mid(TEXT_NO, 5, 1) = "A"
                    DUP_FLG = False
                End If
            
            End If
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 発生元によるデータ確認  2011.03.23



        If Not_SHUSI Then
            NAIGAI = "1"
        Else
            NAIGAI = Soko_T(i, j).NAIGAI
        End If

'        If sts = BtErrKeyNotFound Then
        If Not DUP_FLG Then



''            If Y_GLICS_PUT_PROC(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), NAIGAI, INS_NOW) Then
''                Exit Function
''            End If
            
            
'''''''''''''''''''''2011.03.23 引数追加
'            If Y_GLICS_PUT_PROC(JGYOBU_Code, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME, _
'                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD) Then
                
                
                
            If Y_GLICS_PUT_PROC(JGYOBU_Code, NAIGAI, INS_NOW, _
                                TEXT_NO, _
                                JGYOBU_Code, _
                                CYOK_KBN, _
                                DEN_DT, _
                                IO_KBN, _
                                PM_KBN, _
                                DEN_SYU, _
                                DEN_NO, _
                                CYU_KBN, _
                                HIN_GAI, _
                                HIN_NAI, _
                                HIN_NAME, _
                                YOTEI_QTY, _
                                YOSAN_FROM, _
                                YOSAN_TO, _
                                HOST_SOKO, _
                                HOST_TANA, _
                                SYUK_CODE, _
                                SYUK_NAME, _
                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD, MOTO_TEXT_NO) Then
                
                
                
                
'''''''''''''''''''''2011.03.23 引数追加
                
                Exit Function
            End If

        End If



'-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
    
        Skip_Flg = False
        If Not_SHUSI Then
            Skip_Flg = True
        End If
    
        If PM_KBN = "-" Then
            Skip_Flg = True
        End If
        
        If IO_KBN <> "1" Then
            
            If IO_KBN = "4" And Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" And Trim(HOST_SOKO) = "11B" Then
''''            2011.01.18
''''                If (JGYOBU_Code = SUIHAN Or JGYOBU_Code = DENKA) Then
                If Trim(RYOHEN_TANA) <> "" Then
''''            2011.01.18
                Else
                    Skip_Flg = True
                End If
            Else
                Skip_Flg = True
            End If
        Else
            Skip_Flg = True
        End If
    
    
        WORK_SOKO = KASO_NYUKA_SOKO
    
    
    
''''2011.01.18
''''        Select Case JGYOBU_Code
''''
''''            Case SOJIKI                         '掃除機
''''
''''
''''
''''            Case DENKA, SUIHAN, SENTAKU         '電化、炊飯、洗濯機（アイロン）
''''
''''
''''                Select Case MyCenter
''''
''''                    Case "O"
''''
''''
''''
''''                        '2009.06.01 65番倉庫出力追加
''''                        If (JGYOBU_Code = SUIHAN Or JGYOBU_Code = DENKA) Then
''''                            If IO_KBN = "4" Then
''''                                If Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" Then
''''
''''                                    If Trim(HOST_SOKO) = "11B" Then
''''                                        WORK_SOKO = "65"
''''                                    End If
''''                                End If
''''                            End If
''''                        End If
''''
''''
''''
''''
''''
''''                End Select
''''             Case AIRCON                     'エアコン
''''
''''
''''
''''        End Select
        If Trim(RYOHEN_TANA) <> "" Then
            WORK_SOKO = RYOHEN_TANA
        End If
''''2011.01.18
            
            
            
            
        
    
    
        If Not Skip_Flg Then
                                        
                                        
            
                
                                        '入荷予定重複チェック
            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU_Code)
            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
    
            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
                    Skip_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入荷予定", 0)
                    Exit Function
            End Select
        
            If Not Skip_Flg Then
                                                'トランザクション開始
                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                    Exit Function
                End If
                                            '品目マスタチェック
                If Item_Check_Proc(In_Mode, JGYOBU_Code, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                    GoTo Abort_Tran
                End If
                                            
                                            
                '2012.12.20
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                End If
                '2012.12.20
                                            
                                            '入荷データ作成
                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU_Code)
                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
        
        
                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
        
        
                Last_Proc_F = True              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り
        
        
                '入荷ﾁｪｯｸﾃﾞｰﾀ更新
                Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU_Code)
                Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
    
                WK_Y_QTY = CLng(YOTEI_QTY)
    
                Loop_Cnt = 0
    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                        
                                Loop_Cnt = 0
                        
                                Do
                                
                                    sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                Exit Function
                                            End If
                                            DoEvents
                                            Sleep (500)
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)
                                            Exit Function
                                    End Select
                                
                                Loop
                                WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            Else
                                
                                
                                Loop_Cnt = 0
                                Do
                                    sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                Exit Function
                                            End If
                                            DoEvents
                                            Sleep (500)
                                        
                                                                                    
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)
                                            Exit Function
                                    End Select
                                Loop
                                WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                            End If
                    
                            Exit Do
                        Case BtErrKeyNotFound
                            WK_E_QTY = 0
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                            If ans = vbCancel Then
'                                Exit Function
'                           End If
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                Exit Function
                            End If
                            DoEvents
                            Sleep (500)
                                            
                                            
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)
                            Exit Function
                    End Select
                Loop
                                    '先行入荷数（入荷実績数）
                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
        
                                    '予算単位元
                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                    '予算単位先
                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                    '標準棚番
                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                                    'H倉庫 2006.10.17
                Call UniCode_Conv(Y_NYUREC.H_SOKO, HOST_SOKO)

                                    '入荷リスト出力フラグ   2007.06.12
                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")
            
            '2011.01.18
                Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
                
                com = BtOpGetGreaterEqual
                
                GENSAN_CNT = 0
                
                Do
                    DoEvents
                
                    sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            If StrConv(GENSANREC.JGYOBU, vbUnicode) <> StrConv(Y_NYUREC.JGYOBU, vbUnicode) Or _
                                StrConv(GENSANREC.NAIGAI, vbUnicode) <> StrConv(Y_NYUREC.NAIGAI, vbUnicode) Or _
                                StrConv(GENSANREC.HIN_GAI, vbUnicode) <> StrConv(Y_NYUREC.HIN_NO, vbUnicode) Then
                                Exit Do
                            End If
                        
                        
                            GENSAN_CNT = GENSAN_CNT + 1
                            If GENSAN_CNT > 1 Then
                                GENSANKOKU = ""
                                Exit Do
                            End If
                        
                        
                            GENSANKOKU = StrConv(GENSANREC.GENSANKOKU, vbUnicode)
                        
                        Case BtErrEOF
                            Exit Do
                        
                        Case Else
                            Call File_Error(sts, com, "原産国マスタ", 0)
                            Exit Function
                    End Select
                
                    com = BtOpGetNext
                                
                Loop
                
                
                                
                
                
                Call UniCode_Conv(Y_NYUREC.GENSANKOKU, "")                          '原産国名
                If GENSAN_CNT = 1 Then
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)
                End If
                                                                                    '資材仕入先ﾜｰｸｾﾝﾀｰ
                Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                  '環境種類区分
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)            '環境種類区分適用開始
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)      '環境種類区分数量
                Call UniCode_Conv(Y_NYUREC.ID_NO2, TEXT_NO)                         'ID_NO
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)            '相手先ｺｰﾄﾞ
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                  '受注年月日
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)      '指定納期年月日
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "")                      '入庫関連ﾘｽﾄ出力F
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "")                    '入庫管理ﾘｽﾄ出力F
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "")                    '入庫ﾁｪｯｸﾘｽﾄ出力F
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, WORK_SOKO & "010101")     '入庫棚番
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, "")                       '前借相殺数
                
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2015")       '追加　担当者
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)   '追加　日時
            
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")           '更新　担当者
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")        '更新　日時         2005.11.15
            '2011.01.18

                
                
                
                '2011.03.23 発生元プログラム
                Call UniCode_Conv(Y_NYUREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
                '2011.03.23 元テキスト№
                If Trim(MOTO_TEXT_NO) = "" Then
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, "")
                Else
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
                End If
                
                
                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                
                Loop_Cnt = 0
                
                Do
                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                            If ans = vbCancel Then
'                                Exit Function
'                            End If
                        
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                Exit Function
                            End If
                            DoEvents
                            Sleep (500)
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpInsert, "入荷予定", 0)
                            Exit Function
                    End Select
                Loop
            
'------------ 2005.12.30
'                Select Case JGYOBU_Code
'                    Case AIRCON, SENTAKU
'                        Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
'                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
'                                Exit Function
'                        End Select
'
'                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
'
'                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            MI_QTY = 0
'                        Else
'
'                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                                SUMI_QTY = 0
'                            Else
'                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                                MI_QTY = 0
'                            End If
'                        End If
'
'------------ 2005.12.30
'
'                    Case Else
'
'
'                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            SUMI_QTY = 0
'                        Else
'                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            MI_QTY = 0
'                        End If
'                End Select
                
        
'                Wk_SOKO = KASO_NYUKA_SOKO
'                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'                    Wk_SOKO = KASO_SMODOSHI_SOKO
'
'                End If
        
                
                 MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                 SUMI_QTY = 0
                
                
                '入荷数で在庫データ更新（＋）
'                If Nyuko_Update_Proc(JGYOBU_Code, _
'                                    Soko_T(i, j).NAIGAI, _
'                                    HIN_GAI, _
'                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
'                                    (WORK_SOKO & "01" & "01" & "01"), _
'                                    YOIN_TU_NYUKA, _
'                                    SUMI_QTY, MI_QTY, _
'                                    WS_NO, WS_NO, , _
'                                    DEN_DT & " 伝№:" & DEN_NO, , , , MENU_NO) Then
'                    Exit Function
'
'                End If
            
                
                If Nyuko_Update_Proc(JGYOBU_Code, _
                                    Soko_T(i, j).NAIGAI, _
                                    HIN_GAI, _
                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                    (WORK_SOKO & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, MI_QTY, _
                                    WS_NO, WS_NO, , _
                                    Trim(YOSAN_FROM) & " " & DEN_DT & " 伝№:" & DEN_NO, , , , MENU_NO, , RYOHEN, GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM) Then
                    Exit Function
            
                End If
            
            
            
                '前借り数で在庫データ更新（－）
                If WK_E_QTY <> 0 Then
                '在庫データLOCK
                    If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                        JGYOBU_Code, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        WS_NO, , , 5) Then
                        Exit Function
    
                    End If
        
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                        MI_QTY = WK_E_QTY
                    Else
                        SUMI_QTY = WK_E_QTY
                    End If
            
            
                    If Syuko_Update_Proc(JGYOBU_Code, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        DEN_DT, _
                                        (WORK_SOKO & "01" & "01" & "01"), _
                                        YOIN_MAE_SOUSAI, _
                                        SUMI_QTY, MI_QTY, 0, _
                                        WS_NO, WS_NO, 5) Then
                        Exit Function
        
                    End If
            
            
            
            
            
            
                End If
                
                
                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    GoTo Abort_Tran
                End If
                
                
                Out_Cnt = Out_Cnt + 1
                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                DoEvents
    
            End If
        
        
        
        
        
        
        
        End If
        
        
        
    
    Loop

    New_Nyuka_Update_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


    




End Function
    
Private Function New_Syuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   「出荷予定データ」更新処理
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim Skip_Flg    As Boolean
Dim sts         As Integer
    
Dim ans         As Integer

Dim c               As String * 128

Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
    
Dim INS_NOW         As String * 14
Dim wkCHOKU_KBN     As String * 1

Dim wkSS            As String
        
Dim wkMUKE_CODE     As String

Dim Loop_Cnt        As Integer          '2011.01.15


    
    New_Syuka_Update_Proc = True

    Fast_Flg = True


    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)


    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")

    Do
        Get #FileNo, , New_HS_OUT_SIJREC
        If Left(StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode), 1) < " " Then
            Exit Do
        End If
    
    
        If StrConv(New_HS_OUT_SIJREC.CRLF, vbUnicode) <> vbCrLf Then
'''            Call NG_File_Make_Proc
            Err_FLg = True
            Exit Do
        End If
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode)) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
                                '品目マスタのチェック
'                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Soko_T(i, j).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode))
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Skip_Flg = True
'                        Call Log_Out(LOG_F, "伝票ID=" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                        Exit Function
'                End Select
                
                
                
        If Not Skip_Flg Then
                                                    'トランザクション開始
        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
            Exit Function
        End If
        
        
                                    '品目マスタチェック
        If Item_Check_Proc(Out_Mode, JGYOBU, Soko_T(i, j).NAIGAI, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode)) Then
            GoTo Abort_Tran
        End If
                                                        
'2008.12.16        If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'2008.12.16            IsNumeric(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'2008.12.16        Else
'2008.12.16            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'2008.12.16        End If
                                                        
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            wkCHOKU_KBN = ""
        Else
            wkCHOKU_KBN = "1"
        End If
                                                                                                
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            If Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A1" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A2" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A3" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A4" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A5" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A6" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A7" Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "3")
            End If
        
            If StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000440" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000441" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000442" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000443" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000444" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000445" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000446" Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
            End If
        
        End If
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "3")
        End If
                                                                
        '2006.07.20
        wkMUKE_CODE = StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)
'-----------    2005.12.30
'''                    If JGYOBU = AIRCON Then
'''
'''                        'MTSｺｰﾄﾞの読み替え
'''                        If GetIni(App.EXEName, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
'''                        Else
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, Trim(c))
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                        End If
'''
'''
'''                    End If
            
        'エアコンだった場合向け先に直送先をｾｯﾄ  2004.12.01-->全事業部共通
        
        If (StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode) Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
        Else
            If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) <> 0 Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            End If
        End If


        If Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode)) = "" Then
        
            If GetIni(App.EXEName, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
            Else
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, Trim(c))
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            End If
        End If


        If StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode) = "2" Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "E")
            '貿易は「39040」に集約 2006.05.31
            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, "39040")
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
        
        
        End If
        
        
        
        
        
        
        '部内出荷伝票（ﾃﾞｰﾀ区分＝7、取引区分=29）の向け先変換　2006.06.17
        If Trim(StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode)) = "7" And _
            Trim(StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode)) = "29" Then

            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, _
                Right(StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode), 5))
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")

        End If
        
        
        
        '掃除機の場合は、向け先を戻す。 2007.10.29
        If JGYOBU = SOJIKI Then
        
            If Trim(StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode)) = "7" And _
                Trim(StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode)) = "29" And _
                Trim(StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode)) = "00023210" And _
                Trim(wkMUKE_CODE) = "09002" Then
            
            
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, wkMUKE_CODE)
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            
            
            End If
        
        
        
        
        
        End If
            
        
        
        
        
        '注文区分＝6は２に
''                    If JGYOBU = SENTAKU And StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode) = "S2" Then
''
''                        If StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) = "6" Then
''                            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
''
''
''                        End If
''                    End If

        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "E" Then
            wkCHOKU_KBN = ""
        End If


        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "1" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "2" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "3" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "E" Then
        Else

            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "1")
        End If

'''                    Else
            
            '洗濯機の場合、備考１を直送先にセット 2006.03.25
'''                        If JGYOBU = SENTAKU And StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode) = "S2" Then
'''                            If StrComp(StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode), "FAX", vbTextCompare) Then
'''
'''                                wkSS = ""
'''
'''                                For k = 1 To Len(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
'''                                    If IsNumeric(Mid(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode), k, 1)) Then
'''                                        wkSS = wkSS & Mid(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode), k, 1)
'''                                    Else
'''                                        Exit For
'''                                    End If
'''                                Next k
'''
'''                                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, wkSS)
'''                            End If
'''
'''                        End If
'''
'''
'''
'''                        '他の事業部は現状のまま
'''                        If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'''                            IsNumeric(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'''                        Else
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                        End If
'''                    End If
'                   ↑に移動    2004.12.01
'                    If Len(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'                        IsNumeric(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'                    Else
'                        Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                    End If
                                                        
                                                        
                                                        
                                                        '向け先マスタ読み込み
'                    If Len(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Then
'                        Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))
'                        Call UniCode_Conv(K0_MTS.SS_CODE, "")
            
'                    Else
            
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
'                    End If
                 
                 
                 
                                     
                 
                 
                 
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                
                
'''前事業部ﾏｽﾀ起票 2006.05.31
'''                            If JGYOBU = AIRCON Then
                    'エアコンだった場合向け先に直送先で向け先ﾏｽﾀを新規作成  2004.12.01
                
                    Call UniCode_Conv(MTSREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(MTSREC.DATA_KBN, "")
                    Call UniCode_Conv(MTSREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(MTSREC.SS_CODE, "")
                    Call UniCode_Conv(MTSREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(MTSREC.SS_NAME, "")
                    Call UniCode_Conv(MTSREC.MUKE_DNAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "99")
                    Call UniCode_Conv(MTSREC.FILLER, "")
                    
                    Loop_Cnt = 0
                    
                    Do
                        sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                                                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                                DoEvents
                                Sleep (500)
                                                            
                                                            
                            Case Else
                                Call File_Error(sts, BtOpInsert, "向け先管理ﾏｽﾀ" & "key=" & StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode), 0)
                                GoTo Abort_Tran
                        End Select
                    Loop
                                            
                                            
                                            
                                            
                
'''                            Else
'''                               '他の事業部は現状のまま
'''                                If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                                Else
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                                End If
''                            End If
                
'                          ↑に移動    2004.12.01
'                           If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'                               Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'                               Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                           Else
'                               Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'                               Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                           End If
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ", 0)
                GoTo Abort_Tran
        End Select
                                                        
                                                        
'-----------    2005.12.30
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        '向け先マスタ読み込み
'                    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))
'
'                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                        Case BtErrKeyNotFound
'
'                            If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'                                Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'                                Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                            Else
'                                Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'                                Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                            End If
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
'                            Exit Function
'                    End Select
                        
        
        
'''                    If StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "1" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "2" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "3" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "E" Then
'''
'''                        Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
'''
'''                    End If
                    
                    
                                '出荷予定重複チェック
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    
                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, StrConv(New_HS_OUT_SIJREC.KAIKEI_JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.HOJYO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(New_HS_OUT_SIJREC.TANKA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(New_HS_OUT_SIJREC.ITEM_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, StrConv(New_HS_OUT_SIJREC.ODER_NO_R, vbUnicode))
                    '2011.10.31
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, Left(StrConv(New_HS_OUT_SIJREC.KOSO_KEITAI, vbUnicode), 10))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    
                    
                    
                    If TANA_SPACE Then
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, StrConv(New_HS_OUT_SIJREC.TANABAN1, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(New_HS_OUT_SIJREC.TANABAN2, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, StrConv(New_HS_OUT_SIJREC.TANABAN3, vbUnicode))
                    End If
                    
                    
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, StrConv(New_HS_OUT_SIJREC.ORIGIN1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, StrConv(New_HS_OUT_SIJREC.ORIGIN2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(New_HS_OUT_SIJREC.BIKOU2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_OUT_SIJREC.CHOKU_KBN, vbUnicode))
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, StrConv(New_HS_OUT_SIJREC.UNIT_ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, StrConv(New_HS_OUT_SIJREC.ZAIKO_HIKIATE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, StrConv(New_HS_OUT_SIJREC.GOKON_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, StrConv(New_HS_OUT_SIJREC.JYUCHU_ZAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, StrConv(New_HS_OUT_SIJREC.KYOKYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHOHIN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.S_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.S_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(New_HS_OUT_SIJREC.CHOHA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, StrConv(New_HS_OUT_SIJREC.JYU_HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, StrConv(New_HS_OUT_SIJREC.HIN_CHANGE_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, StrConv(New_HS_OUT_SIJREC.MODULE_EXCHANGE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAIKO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, StrConv(New_HS_OUT_SIJREC.NOUKI_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, StrConv(New_HS_OUT_SIJREC.SERVICE_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(New_HS_OUT_SIJREC.KISHU_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, StrConv(New_HS_OUT_SIJREC.ENVIRONMENT_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, StrConv(New_HS_OUT_SIJREC.KEPIN_KAIJYO, vbUnicode))
                    
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    
'2008.11.28                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    Call UniCode_Conv(Y_SYUREC.UPD_NOW, INS_NOW)            '2008.11.28
                
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                    
                    sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                        GoTo Abort_Tran
                    End If
                                
                                
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
                                
                                
                                
                    
                    Call LOG_OUT(LOG_F, "Y_SYUKA.DAT DUP 事業部=" & JGYOBU & "伝票ＩＤ＝" & StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Skip_Flg = True
                
                
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
                        Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")
                        Write #DUP_SYUKANo, "出荷日", "伝票№", "支払先ｺｰﾄﾞ", "倉庫/ＳＳｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"
                        Fast_Flg = False
                    End If
                
                
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode)
                
                
                
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    
                    
                                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    
    '''2006.07.15                    If (JGYOBU = "D" Or JGYOBU = "4") And StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) = "E" Then
    '''2006.07.15                        Call UniCode_Conv(Y_SYUREC.NAIGAI, "2")
    '''2006.07.15                    Else
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
    '''2006.07.15                    End If
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, StrConv(New_HS_OUT_SIJREC.KAIKEI_JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode))
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.HOJYO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(New_HS_OUT_SIJREC.TANKA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(New_HS_OUT_SIJREC.ITEM_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, StrConv(New_HS_OUT_SIJREC.ODER_NO_R, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, StrConv(New_HS_OUT_SIJREC.KOSO_KEITAI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    If TANA_SPACE Then
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, StrConv(New_HS_OUT_SIJREC.TANABAN1, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(New_HS_OUT_SIJREC.TANABAN2, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, StrConv(New_HS_OUT_SIJREC.TANABAN3, vbUnicode))
                    End If
                    
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, StrConv(New_HS_OUT_SIJREC.ORIGIN1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, StrConv(New_HS_OUT_SIJREC.ORIGIN2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(New_HS_OUT_SIJREC.BIKOU2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode))
                    
                    '2006.07.31 項目内容をそのままｾｯﾄ
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_OUT_SIJREC.CHOKU_KBN, vbUnicode))
    '''                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, wkCHOKU_KBN)
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, StrConv(New_HS_OUT_SIJREC.UNIT_ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, StrConv(New_HS_OUT_SIJREC.ZAIKO_HIKIATE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, StrConv(New_HS_OUT_SIJREC.GOKON_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, StrConv(New_HS_OUT_SIJREC.JYUCHU_ZAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, StrConv(New_HS_OUT_SIJREC.KYOKYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHOHIN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.S_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.S_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(New_HS_OUT_SIJREC.CHOHA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, StrConv(New_HS_OUT_SIJREC.JYU_HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, StrConv(New_HS_OUT_SIJREC.HIN_CHANGE_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, StrConv(New_HS_OUT_SIJREC.MODULE_EXCHANGE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAIKO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, StrConv(New_HS_OUT_SIJREC.NOUKI_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, StrConv(New_HS_OUT_SIJREC.SERVICE_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(New_HS_OUT_SIJREC.KISHU_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, StrConv(New_HS_OUT_SIJREC.ENVIRONMENT_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, StrConv(New_HS_OUT_SIJREC.KEPIN_KAIJYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                    
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    
                    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
                    
                    
                    Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")              '2006.09.07
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")            '2006.09.07
                    
                    
                    Call UniCode_Conv(Y_SYUREC.UPD_NOW, "")                 '2008.11.28
                    
                    
                    Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
            
    
                    
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    GoTo Abort_Tran
                                End If
                                
                                DoEvents
                                Sleep (500)
                            Case Else
                                Call File_Error(sts, BtOpInsert, "出荷予定", 0)
                                GoTo Abort_Tran
                        End Select
                    Loop
    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
        
        
        
        
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                    DoEvents
        
        
        
        
        
        
        
                    If SYUKA_LOG_ON Then
                        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "出荷予定", 0)
                    GoTo Abort_Tran
            End Select
                    
                                
                    
                    
                    
                    
                    
                    
                    
                    
                    
        End If
    Loop
    
        
    Close #DUP_SYUKANo
        
        
        
    New_Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
                                            'ヘッダー印刷（「品番変更リスト」）
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "＊＊＊　品番変更リスト　＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print "------- 品番（外部）-------";
    Printer.Print Tab(30);
    Printer.Print "------- 品番（内部）-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "伝票日付";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "入出庫区";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "入出数";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "倉";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "注文区";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "出荷先"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '明細印刷（「品番変更リスト」）
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = 0 To LBox_Hin.ListCount - 1
        
'        Ldata = LBox_Hin.List(i)
'
'                                        'ヘッダーコントロール
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If
'
'                                        '明細印刷
'        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '事業部，ﾃｷｽﾄ№，国内外　除外'
'
'        Printer.Print Tab(MGN_L);
'        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（外部）
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '外部品番変更？
'            Printer.Print Tab(MGN_L + 15);
'            Printer.Print Work;                                     'マスタ品番（外部）
'        End If
'
'        Printer.Print Tab(MGN_L + 30);
'        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（内部）
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "0" Then                               '内部品番変更？
'            Printer.Print Tab(MGN_L + 44);
'            Printer.Print Work;                                     'マスタ品番（内部）
'        End If
'
'        Printer.Print Tab(MGN_L + 58);                              '伝票日付
'        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);
'
'        Printer.Print Tab(MGN_L + 69);                              '入出庫区分
'        wk_IO = ChrCut(Ldata, 1)
'        Select Case wk_IO
'            Case IO_KBN_URI
'                Printer.Print wk_IO & " " & (IO_KBN_0);
'            Case IO_KBN_NYU
'                Printer.Print wk_IO & " " & (IO_KBN_1);
'            Case IO_KBN_SYU
'                Printer.Print wk_IO & " " & (IO_KBN_2);
'            Case IO_KBN_ZAT
'                Printer.Print wk_IO & " " & (IO_KBN_3);
'            Case Else
'                Printer.Print wk_IO;
'        End Select
'
'        Printer.Print Tab(MGN_L + 78);
'        Printer.Print ChrCut(Ldata, 6);                             '伝票№
'
'        Printer.Print Tab(MGN_L + 85);                              '入出庫数
'        Wqty = CLng(ChrCut(Ldata, 6))
'
'
'        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
'
'        Printer.Print Work;
'
'        Printer.Print Tab(MGN_L + 93);
'        Printer.Print ChrCut(Ldata, 2);                             '倉庫区分（ﾎｽﾄ）
'
'        Printer.Print Tab(MGN_L + 96);                              '注文区分
'        Select Case Left(Ldata, 1)
'            Case CYU_KBN_TUK
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
'            Case CYU_KBN_SPO
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
'            Case CYU_KBN_HJU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
'            Case CYU_KBN_BOU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
'            Case Else
'                Printer.Print ChrCut(Ldata, 1);
'        End Select
'
'        Printer.Print Tab(MGN_L + 103);
'        Printer.Print ChrCut(Ldata, 5);                             '支給先／出荷先7
'
'        Printer.Print Tab(MGN_L + 110);                             '変更メッセージ
'        Select Case Left(Ldata, 1)
'            Case "0"
'                Printer.Print "内部変更 ﾏｽﾀ品番入替";
'            Case "1"
'                Printer.Print "外部変更 ﾏｽﾀ品番入替";
'            Case "2"
'                Printer.Print "在庫有！外部変更不可";
'        End Select
        
        Printer.Print LBox_Hin.List(i)
        
        Call LOG_OUT(LOG_F, LBox_Hin.List(i))
        
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    End If

End Sub

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String


    Call NG_File_Make_Proc

    Err_FLg = False

    '---------------------------------------------  事業部毎メインループ
    For i = 0 To UBound(JGYOBU_T)
        
        In_Cnt = 0
        Out_Cnt = 0

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents

'2007.06.22　入庫取り込みを復活

        FileNo = FreeFile
        FileName = New_HS_IN_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

        On Error GoTo Error_Proc

        Open FileName For Input As #FileNo

        On Error GoTo 0


        If New_Nyuka_Update_Proc(JGYOBU_T(i).CODE) Then     '入荷予定データ更新処理

            Unload Me

        End If


        Close #FileNo

        '-----------------------------------------------
    
        FileNo = FreeFile
        FileName = New_HS_OUT_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
        
        On Error GoTo Error_Proc
        
        Open FileName For Binary As #FileNo
    
        On Error GoTo 0
    
    
        If New_Syuka_Update_Proc(JGYOBU_T(i).CODE) Then  '出荷予定データ更新処理

            Unload Me
        End If
    
    
        Close #FileNo
    
    
    
    
    Next i


    If Not Err_FLg Then
        Call NG_File_Kill_Proc
    End If


    Unload Me

Error_Proc:

    Unload Me


End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer


Dim sBuffer     As String * 255
Dim com         As String
    
Dim Max_Soko    As Integer
    
    If App.PrevInstance Then
 '       Beep
 '       MsgBox "同一プログラム実行中です。"
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
                                 '出荷重複データ出力ファイル名取り込み
    If GetIni("FILE", "DUP_SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "出荷重複データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    DUP_SYUKA_DATA = Trim(c)
                               
    If JGYOB_TB_Set(1) Then      '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                '倉庫最大数を取り込み
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                    
                                
                                
                                
                                '在庫取り込み用テーブル作成
    ReDim Soko_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '倉庫情報取り込み
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '有効倉庫獲得
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "倉庫情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
    
            If Trim(c) = "**" Then  '倉庫指定終了
                Exit Do
            End If
    
    
'            ReDim Preserve JSOKO_T(i).JSOKO_T(0 To j)
            Soko_T(i, j).HS_SOKO = Trim(c)
                                '国内外情報獲得
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "国内外情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
            
            Soko_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                
                                
                                
                                
    '良品返品入庫棚番   2011.01.18
    If GetIni(App.EXEName, "RYOHEN_TANA", App.EXEName, c) Then
        RYOHEN_TANA = ""
    Else
        RYOHEN_TANA = RTrim(c)
    End If
                                
                                
                                '品名による除外 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                '品名による除外 2011.07.04
                                
                                
                                '入庫データファイル名の獲得
    If GetIni("FILE", "NEW_HS_SIJ_IN", "SYS", c) Then
        Beep
        MsgBox "入庫データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    New_HS_IN_SIJ = Trim(c)
                                
                                
                                '新ﾚｲｱｳﾄ 出庫データファイル名の獲得 2006.05.23
    If GetIni("FILE", "NEW_HS_SIJ_OUT", "SYS", c) Then
        Beep
        MsgBox "新　出庫データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    New_HS_OUT_SIJ = Trim(c)
                                
                                
                                
                                '「通常入荷」要因の獲得
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Beep
        MsgBox "「通常入荷」要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                '「前借相殺」要因の獲得
    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
        Beep
        MsgBox "「前借相殺」要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_MAE_SOUSAI = Trim(c)
                                
                                '仮想入荷倉庫の獲得
    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "仮想入荷倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KASO_NYUKA_SOKO = Trim(c)
                                '仮想支給戻し倉庫の獲得
    If GetIni("SYSTEM", "KASO_SMODOSHI ", "SYS", c) Then
        Beep
        MsgBox "仮想入荷倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KASO_SMODOSHI_SOKO = Trim(c)
                                
                                
                                'その他向け先（国内）の獲得
    If GetIni(App.EXEName, "ETC_MTS_NAI", App.EXEName, c) Then
        Beep
        MsgBox "その他向け先の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ETC_MTS_NAI = Trim(c)
                                
                                'その他向け先（海外）の獲得
    If GetIni(App.EXEName, "ETC_MTS_GAI", App.EXEName, c) Then
        Beep
        MsgBox "その他向け先の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ETC_MTS_GAI = Trim(c)
                                
'---------------------------------------------- 'ﾒﾆｭｰ№の獲得    2007.11.06
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = RTrim(c)
    End If
                                
                                '洗濯機専用
    If GetIni(App.EXEName, "CENTER", "SYS", c) Then
        MyCenter = "O"
    Else
        MyCenter = Trim(c)
    End If
'---------------------------------------------- '良品返品の要因 2009.07.10
    RYOHEN = YOIN_TU_NYUKA
    If GetIni(App.EXEName, "RYOHEN", App.EXEName, c) Then
    Else
        RYOHEN = RTrim(c)
    End If
                                
                                
                                'その他直送先の獲得
'    If GetIni(App.EXEName, "ETC_SS_NAI", "SYS", c) Then
'        Beep
'        MsgBox "その他直送先の獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
'    ETC_SS_NAI = Trim(c)
                                
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

'---------------------------------------------- '棚番設定情報の獲得    2009.03.07
    If GetIni(App.EXEName, "TANA_SPACE", App.EXEName, c) Then
        TANA_SPACE = False
    Else
        If Trim(c) = "1" Then
            TANA_SPACE = True
        Else
            TANA_SPACE = False
        End If
    End If

'---------------------------------------------- '商品化ﾃﾞﾌｫﾙﾄ    2012.12.20
    If GetIni(App.EXEName, "GOODS_F", App.EXEName, c) Then
        GOODS_F = "0"
    Else
        If Trim(c) = "1" Then
            GOODS_F = "1"
        Else
            GOODS_F = "0"
        End If
    End If
'---------------------------------------------- '商品化ﾃﾞﾌｫﾙﾄ    2012.12.20



                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタ（更新用ワーク）ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
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
                                'コードマスタＯＰＥＮ   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '原産国マスタＯＰＥＮ   2010.07.08
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                'PNマスタＯＰＥＮ   2010.09.01
    If PN_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    If Country_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷ﾁｪｯｸﾃﾞｰﾀＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '照合用入荷予定ＯＰＥＮ 2007.06.15
    If Y_GLICS_Open(BtOpenNomal) Then
        Unload Me
    End If

    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If

'発番マスタＯＰＥＮ ################################################################## 2005/05/16 Add ↓
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
'#################################################################################### 2005/05/16 Add ↑
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '原産国マスタＯＰＥＮ 2011.01.18
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020151.FontName
        .Size = F1020151.FontSize
    End With
    Set Printer.Font = NormalFont

    Last_Proc_F = False         '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグクリア


    '仕向け先獲得       2005.12.30
    i = -1
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    com = BtOpGetGreater
    SHIMUKE_Flg = False
    
    Do
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SHIMUKE_T(0 To i)
            
            
                SHIMUKE_Flg = True
            
                SHIMUKE_T(i).SHIMUKE_CODE = StrConv(P_CODEREC.C_Code, vbUnicode)
                SHIMUKE_T(i).JGYOBU = StrConv(P_CODEREC.OPTION1, vbUnicode)
                SHIMUKE_T(i).NAIGAI = StrConv(P_CODEREC.OPTION2, vbUnicode)
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ", 0)
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
    
    
    '仕向け先獲得       2005.12.30


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り？
'        Call Last_Proc
'    End If

                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '品目マスタ（更新用ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
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
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
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
                                            '入荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定")
        End If
    End If
                                            '照合用入荷予定ＣＬＯＳＥ   2007.06.16
    sts = BTRV(BtOpClose, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "照合用入荷予定")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            '入荷ﾁｪｯｸﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷ﾁｪｯｸﾃﾞｰﾀ")
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020151 = Nothing

    End
End Sub
Private Function Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ」チェック＆更新処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    
Dim Loop_Cnt    As Integer          '2011.01.19
    
    
    Item_Check_Proc = True

    HIN_CHANGE = 0
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                


                    If Mode = In_Mode Then          '対内品番変更のチェック
    '                Else
    
                        If Len(Trim(HIN_NAI)) <> 0 Then
                            If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                                HIN_CHANGE = NAI_CHANGE
                                BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '内部品番入れ替え
                                Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                            
                            
                                '担当者更新追加 2009.11.11
                                    
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            
                            
                            
                            
                            
                            
                            End If
                        End If
                    
                    
                    
                                        
                        '---------------    2010.07.08  ▽
                        '原産国入れ替えチェック
                        If Len(Trim(GENSANKOKU)) <> 0 Or Len(Trim(GEN_GENSANKOKU)) <> 0 Or Len(Trim(SHIIRE_WORK_CENTER)) <> 0 Then
    '                        If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> Trim(GENSANKOKU) Then
                                '原産国入れ替え
                                
                            
                                If Trim(GENSANKOKU) <> "" Then
                                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                                Else
                                    Debug.Print
                                End If
                                
                                
    '                            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
    '
    '                                Call UniCode_Conv(ITEMREC.GENSANKOKU, GENSANKOKU)
    '
    '
    '                            End If
                                
                                
                                Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                                Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                            
                            
                                '担当者更新追加 2009.11.11
                                    
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                
                                
                            
    '                        End If
                        End If
                        '---------------    2010.07.08  △
                    
                    
                        '---------------    2010.07.27  ▽
                        '環境区分チェック
                        If Len(Trim(KANKYO_KBN)) <> 0 Or Len(Trim(KANKYO_KBN_ST)) <> 0 Or Len(Trim(KANKYO_KBN_SURYO)) <> 0 Then
                            
                            
                            
                            If Val(KANKYO_KBN_SURYO) = 0 Then
                            Else
                            
                                
                                '環境区分入れ替え
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                                    
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                                
                                
                                '担当者更新追加 2009.11.11
                                        
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            End If
                                
                            
                        End If
                        '---------------    2010.07.08  △
                    
                    
                    
                    End If
                
                
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '事業部
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '品番（外部）
                                                            '品名
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '新規品番時*をセット2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '品番（内部）
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '品番（内部）
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '資材コード
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '排他フラグ
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '使用子機ＩＤ
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '使用中プログラム
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '元事事業部
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '商品化有無
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 追加(業務管理項目) ▽
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '業務管理　 仕入区分
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           販売区分
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           収支単位
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           組立製品
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           標準粗利売価単価　9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           標準粗利売価設定日
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           標準粗利原価単価  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           標準粗利原価設定日
                                            
                                            
                                                                        '           仕入先情報
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '仕入単価
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ﾛｯﾄ数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ﾘｰﾄﾞﾀｲﾑ
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           前月在庫金額
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           資材区分
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ラベル貼付
'*------------------------------------------ 2005.11.15 追加(業務管理項目) △

'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
                
                
                                                                '           価格(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
                
                
                
'*------------------------------------------ 2008.08.26 新規追加項目一式 ▽
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '在庫管理対象有無 1:対象 0:対象外
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           前月在庫数量
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           最終出荷数
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS在庫(S2) 袋井用
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS在庫(P2) 袋井用
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      'ﾕﾆｯﾄ部品区分
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '国内供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '海外供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '標準単価   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '最終仕入先コード   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '最終仕入単価       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'ﾒｰｶｰ名称           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '再梱包ﾏｰｸ          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '形式               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '材質               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '印刷する／しない   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '商品化　工数       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '商品化　工数原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '商品化　工数売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '商品化　単価設定日 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '商品化　資材原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '商品化　資材売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '商品化　単価設定日 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '使用テープ種類     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '使用テープ長       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '棚番マーク         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '請求単価　メモ     2008.04.15
    
                '2010.07.08 ▽
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '原産国             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '原産国
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '原産国
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 △
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY")
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '外装単価 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '生産ロット         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '分レート           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '集合梱包           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '単価設定担当者     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '仕向け先           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '請求区分           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            'ラベル貼り枚数     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '資材件数     　    2008.08.20追加
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '同梱件数           2008.08.20追加
         

'*------------------------------------------ 2008.08.26 新規追加項目一式 △
                
                
                                
                
                
                
                '↓2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '↑2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '状態区分           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '検品ﾒｯｾｰｼﾞ 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '請求Ｆ 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '追加担当者
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '追加日時
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '更新担当者
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '更新日時
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
           
            
            Case BtErrDEAD_LOCK
                Exit Function
                        
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                Call File_Error(sts, com, "品目マスタ", 0)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '構成マスタの追加       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '仕向け先ｺｰﾄﾞ
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '事業部
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '国内外
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '品番
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            'ﾃﾞｰﾀ区分
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '追番
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '基本クラス
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '備考
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '更新担当者
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '更新日時
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            Case BtErrDEAD_LOCK
                                Exit Function
                            Case Else
                                Call File_Error(sts, BtOpInsert, "構成マスタ", 0)
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    Item_Check_Proc = False

End Function

Private Function old_Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ」チェック＆更新処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim Loop_Cnt        As Integer          '2011.01.15
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    old_Item_Check_Proc = True

    HIN_CHANGE = 0
    
    
           
'    If Mode = Out_Mode Then
'        Item_Check_Proc = False
'        Exit Function
'    End If
    
'    If Len(Trim(StrConv(HS_IN_SIJREC.HIN_NAI, vbUnicode))) <> 0 Then
'
'        Call UniCode_Conv(K2_ITEM.JGYOBU, JGYOBU)
'        Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
'        Call UniCode_Conv(K2_ITEM.HIN_NAI, StrConv(HS_IN_SIJREC.HIN_NAI, vbUnicode))
'
'
'        Do                          '対外品番入れ替えのループ
'
'            sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'            Select Case sts
'                Case BtNoErr
'                    If StrConv(ITEMREC.HIN_GAI, vbUnicode) <> Trim(HIN_GAI) Then
'                                    '外部品番の入れ替えの為の在庫有無チェック
'                        If Zaiko_Syukei_Proc(Sumi_Qty, MI_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
'                            Exit Function
'                        End If
'
'
'                        If (Sumi_Qty + MI_Qty) = 0 Then
'                            Do
'                                sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                                Select Case sts
'                                    Case BtNoErr
'                                        Exit Do
'
'
'                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                        If ans = vbCancel Then
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
'                                        Exit Function
'                                End Select
'                            Loop
'
'
'                            HIN_CHANGE = GAI_CHANGE
'                            BEF_GAI = HIN_GAI
'
'                        Else
'                            '品番入れ替え不可
'                            sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                            If sts Then
'                                Call File_Error(sts, BtOpUnlock, "品目マスタ")
'                                Exit Function
'                            End If
'                            HIN_CHANGE = NOT_GAI_CHANGE
'                            BEF_GAI = HIN_GAI
'
'                        End If
'                    End If
'
'                    Exit Do
'
'                Case BtErrKeyNotFound
'                    Exit Do
'                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                    Beep
'                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                    If ans = vbCancel Then
'                        Exit Function
'                    End If
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
'                    Exit Function
'            End Select
'
'        Loop
'
'    End If

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr

                If Mode = In_Mode Then          '対内品番変更のチェック
'                Else

                    If Len(Trim(Trim(HIN_NAI))) <> 0 Then
                        If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                            HIN_CHANGE = NAI_CHANGE
                            BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                            '内部品番入れ替え
                            Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                        
                        
                            '担当者更新追加 2009.11.11
                                
                                                                                    '更新担当者
                            Call UniCode_Conv(ITEMREC.UPD_TANTO, "2015")
                                                                                    '更新日時
                            Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                        
                        
                        End If
                    End If
                End If
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 差し替え    2012.07.05
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '事業部
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '品番（外部）
                                                            '品名
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '新規品番時*をセット2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '品番（内部）
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '品番（内部）
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '資材コード
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '排他フラグ
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '使用子機ＩＤ
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '使用中プログラム
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '元事事業部
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '商品化有無
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 追加(業務管理項目) ▽
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '業務管理　 仕入区分
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           販売区分
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           収支単位
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           組立製品
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           標準粗利売価単価　9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           標準粗利売価設定日
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           標準粗利原価単価  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           標準粗利原価設定日
                                            
                                            
                                                                        '           仕入先情報
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '仕入単価
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ﾛｯﾄ数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ﾘｰﾄﾞﾀｲﾑ
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           前月在庫金額
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           資材区分
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ラベル貼付
'*------------------------------------------ 2005.11.15 追加(業務管理項目) △

'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
                
                
                                                                '           価格(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
                
                
                
'*------------------------------------------ 2008.08.26 新規追加項目一式 ▽
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '在庫管理対象有無 1:対象 0:対象外
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           前月在庫数量
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           最終出荷数
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS在庫(S2) 袋井用
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS在庫(P2) 袋井用
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      'ﾕﾆｯﾄ部品区分
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '国内供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '海外供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '標準単価   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '最終仕入先コード   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '最終仕入単価       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'ﾒｰｶｰ名称           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '再梱包ﾏｰｸ          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '形式               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '材質               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '印刷する／しない   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '商品化　工数       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '商品化　工数原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '商品化　工数売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '商品化　単価設定日 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '商品化　資材原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '商品化　資材売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '商品化　単価設定日 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '使用テープ種類     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '使用テープ長       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '棚番マーク         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '請求単価　メモ     2008.04.15
    
                '2010.07.08 ▽
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '原産国             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '原産国
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '原産国
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 △
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY")
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '外装単価 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '生産ロット         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '分レート           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '集合梱包           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '単価設定担当者     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '仕向け先           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '請求区分           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            'ラベル貼り枚数     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '資材件数     　    2008.08.20追加
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '同梱件数           2008.08.20追加
         

'*------------------------------------------ 2008.08.26 新規追加項目一式 △
                
                
                                
                
                
                
                '↓2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '↑2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '状態区分           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '検品ﾒｯｾｰｼﾞ 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '請求Ｆ 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '追加担当者
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '追加日時
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '更新担当者
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '更新日時
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do




                
                
'                com = BtOpInsert
'
'                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '事業部
'                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '国内外
'                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '品番（外部）
'
'                If Mode = In_Mode Then                      '品名
'                    Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
'                Else
'                    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
'                End If
'
''                If Mode = Out_Mode Then                      '標準棚番
''                    If Len(Trim(HS_IN_SIJREC.TANABAN1)) = 0 Then
''                        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
''                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
''                        Call UniCode_Conv(ITEMREC.ST_RETU, "")
''                        Call UniCode_Conv(ITEMREC.ST_REN, "")
''                        Call UniCode_Conv(ITEMREC.ST_DAN, "")
''                    Else
''                        Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
''                        Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 1, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_RETU, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 3, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_REN, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 5, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_DAN, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 7, 2)))
''
''                    End If
''                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
''2009.01.21 "**" に変更                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
''                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
''                    Call UniCode_Conv(ITEMREC.ST_REN, "")
''                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
'
''                End If
'
'                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
'                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
'                Call UniCode_Conv(ITEMREC.BEF_REN, "")
'                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
'
'                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
'                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
'
'                If Mode = In_Mode Then          '対内品番
'                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
'                Else
'                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
'                End If
'
'                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
'                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
''                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '資材コード
'                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
'                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
'                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
'                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
'
''                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '排他フラグ
''                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '使用子機ＩＤ
''                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '使用中プログラム
'
'                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
'                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
''                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '元事事業部
'                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
'                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
'                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
'                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
'                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '商品化有無
'                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
'                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
'                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
'                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
'
'
'
'
'
''*------------------------------------------ 2005.11.15 追加(業務管理項目) ▽
'                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '業務管理　 仕入区分
'                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           販売区分
'                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           収支単位
'                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           組立製品
'                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           標準粗利売価単価　9(8)V99
'                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           標準粗利売価設定日
'                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           標準粗利原価単価  9(8)V99
'                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           標準粗利原価設定日
'
'
'                                                                        '           仕入先情報
'                For i = 0 To 2
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '仕入単価
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ﾛｯﾄ数
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ﾘｰﾄﾞﾀｲﾑ
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ﾘｰﾄﾞﾀｲﾑ
'
'                Next i
'
'                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           前月在庫金額
'                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           資材区分
'                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ラベル貼付
''*------------------------------------------ 2005.11.15 追加(業務管理項目) △
'
''*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
'                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
'                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
'                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
'                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
'                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
'                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
'                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
'                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
'                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
'                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
'                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
'                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
'                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
'                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
'                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
'                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
'                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
'
'
'
'
'
''*------------------------------------------ 2008.08.26 新規追加項目一式 ▽
'
'                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
'                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '在庫管理対象有無 1:対象 0:対象外
'
'                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
'
'                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           前月在庫数量
'                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           最終出荷数
'
'                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS在庫(S2) 袋井用
'                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS在庫(P2) 袋井用
'
'                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
'
'
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価   2006.07.28
'
'                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '最終仕入先コード   2007.05.29
'                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '最終仕入単価       2007.05.29
'
'                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
'                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'ﾒｰｶｰ名称           2007.06.06
'
'
'                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '再梱包ﾏｰｸ          2007.11.08
'
'                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2008.02.14
'
'                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '形式               2008.02.14
'                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '材質               2008.02.14
'                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
'                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
'                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
'
'                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '印刷する／しない   2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '商品化　工数       2008.02.14
'
'                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '商品化　工数原価   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '商品化　工数売価   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '商品化　単価設定日 2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '商品化　資材原価   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '商品化　資材売価   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '商品化　単価設定日 2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
'
'                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '使用テープ種類     2008.02.14
'                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '使用テープ長       2008.02.14
'
'                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '棚番マーク         2008.04.02
'
'
'                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '請求単価　メモ     2008.04.15
'
'
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '原産国             2008.06.11
'
'
'
'                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '外装単価 9(8)V99   2008.06.12
'                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2008.06.12
'                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2008.06.12
'
'
'                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '生産ロット         2008.07.07
'                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '分レート           2008.07.07
'                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '集合梱包           2008.07.07
'
'
'                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '単価設定担当者     2008.07.09
'
'                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '仕向け先           2008.07.09
'
'                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '請求区分           2008.07.16
'
'                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            'ラベル貼り枚数     2008.07.19
'
'                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '資材件数     　    2008.08.20追加
'                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '同梱件数           2008.08.20追加
'
'*------------------------------------------ 2008.08.26 新規追加項目一式 △
'
'                '↓2009.02.20
'                For i = 0 To 9
'                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
'                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
'                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")
'
'                Next i
'
'
'                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
'                '↑2009.02.20
'
'                Call UniCode_Conv(ITEMREC.STAT, "1")                    '状態区分           2009.01.21
'
'
'
'
'''''''''''''''' 2011.07.05  '''''''''''''''''''''''''''''''''''''''''''''
'
'
'                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '検品ﾒｯｾｰｼﾞ 2009.08.28
'                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '請求Ｆ 2009.08.28
'
'                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
'
'                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
'                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
'                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
'                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
'                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
'
'
'
'
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
'
'
'                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
'
'                    If Val(KANKYO_KBN_SURYO) <> 0 Then
'
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
'
'                    End If
'
'                End If
'
'
'                                                                        '追加担当者
'                '2010.09.01
''                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
'                If Mode = Out_Mode Then
'                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
'                Else
'                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
'                End If
'                '2010.09.01
'
'                                                                        '追加日時
'                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'
'
'
'
'
'                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
'                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
'
'
'
'                Call UniCode_Conv(ITEMREC.BIKOU20, "")
'
'                '2011.07.05
'                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
'
'                If NOT_Hin_Name_F Then
'                    For i = 0 To UBound(NOT_Hin_Name)
'                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
'                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
'                            Exit For
'                        End If
'                    Next i
'                End If
'                '2011.07.05
'
'
'
'''''''''''''''' 2011.07.05  '''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
'
'
'
'
'
'''''''''''''''' 2011.07.05  ''''''''''''''''''''''''''''''''''''''''''''' delete
''                sBuffer = Space(255)                                    '2009.01.21
''                If GetComputerNameA(sBuffer, 255) <> 0 Then
''                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
''                Else
''                    wkTanto = "???"
''                End If
''
''
''                                                                        '追加担当者
''                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
''                                                                        '追加日時
''                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'''''''''''''''' 2011.07.05  ''''''''''''''''''''''''''''''''''''''''''''' delete
'
'
'
'
'
'
'
'
'
'
'
'                Call UniCode_Conv(ITEMREC.FILLER, "")
'                                                                        '更新担当者
'                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
'                                                                        '更新日時
'                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'
'
'
'
'
'                Exit Do
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 差し替え    2012.07.05
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If


                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)


            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)
            
            Case Else
                Call File_Error(sts, com, "品目マスタ", 0)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '構成マスタの追加       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '仕向け先ｺｰﾄﾞ
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '事業部
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '国内外
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '品番
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            'ﾃﾞｰﾀ区分
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '追番
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '基本クラス
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '備考
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '更新担当者
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '更新日時
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                                DoEvents
                                Sleep (500)
                            
                            Case Else
                                Call File_Error(sts, BtOpInsert, "構成マスタ", 0)
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    old_Item_Check_Proc = False

End Function

Sub NG_File_Make_Proc()
'----------------------------------------------------------------------------
'                   異常終了ファイル出力処理
'----------------------------------------------------------------------------
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "異常終了ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    stream = FreeFile
    Open NG_FILE For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Print #stream, Buf
    Close stream



End Sub

Sub NG_File_Kill_Proc()
'----------------------------------------------------------------------------
'                   異常終了ファイル削除処理
'----------------------------------------------------------------------------
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "異常終了ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    On Error GoTo Err_Proc
    Kill (NG_FILE)

Err_Proc:

End Sub

'Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String) As Integer
Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String, _
                                                                        TEXT_NO As String, _
                                                                        JGYOBU_Code As String, _
                                                                        CYOK_KBN As String, _
                                                                        DEN_DT As String, _
                                                                        IO_KBN As String, _
                                                                        PM_KBN As String, _
                                                                        DEN_SYU As String, _
                                                                        DEN_NO As String, _
                                                                        CYU_KBN As String, _
                                                                        HIN_GAI As String, _
                                                                        HIN_NAI As String, _
                                                                        HIN_NAME As String, _
                                                                        YOTEI_QTY As String, _
                                                                        YOSAN_FROM As String, _
                                                                        YOSAN_TO As String, _
                                                                        HOST_SOKO As String, _
                                                                        HOST_TANA As String, _
                                                                        SYUK_CODE As String, _
                                                                        SYUK_NAME As String, _
                                                                        GENSANKOKU As String, GEN_GENSANKOKU As String, SHIIRE_WORK_CENTER As String, KANKYO_KBN As String, KANKYO_KBN_ST As String, KANKYO_KBN_SURYO As String, ID_NO2 As String, AITESAKI_CODE As String, JYUCHU_YMD As String, SHITEI_NOUKI_YMD As String, MOTO_TEXT_NO As String) As Integer
'----------------------------------------------------------------------------
'           照合用入荷予定ファイル出力処理
'           2007.06.15
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
        
    
Dim Loop_Cnt        As Integer          '2011.01.15
    
    Y_GLICS_PUT_PROC = True
        
    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, TEXT_NO)


    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
    
    Call UniCode_Conv(Y_GLICSREC.HIN_NO, HIN_GAI)
    Call UniCode_Conv(Y_GLICSREC.DEN_NO, DEN_NO)
    
    
    '2008.01.10 マイナスの対応
    If YOTEI_QTY >= 0 Then
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
    Else
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "000000"))
    End If
    
    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, CYU_KBN)
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, HIN_NAME)
    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
    
    
    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))


                        '先行入荷数（入荷実績数）
    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")

                        '予算単位元
    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, YOSAN_FROM)
                        '予算単位先
    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, YOSAN_TO)
                        '標準棚番
    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, HIN_NAI)
                        'H倉庫 2006.10.17
    Call UniCode_Conv(Y_GLICSREC.H_SOKO, HOST_SOKO)

                        '入荷リスト出力フラグ   2007.06.12
    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
                        '直送区分
    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, CYOK_KBN)
                        '入出庫区分
    Call UniCode_Conv(Y_GLICSREC.IO_KBN, IO_KBN)
                        '赤黒区分
    Call UniCode_Conv(Y_GLICSREC.PM_KBN, PM_KBN)
                        '伝票種別
    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, DEN_SYU)
                        '支給先／出荷先
    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, SYUK_CODE)
                        '支給先／出荷先名
    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, SYUK_NAME)
                        '挿入年月日
    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
    
    
    '----------------   2010.07.08 ▽
    Call UniCode_Conv(Y_GLICSREC.GENSANKOKU, GENSANKOKU)                    '原産国名
    Call UniCode_Conv(Y_GLICSREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '現物表示原産国名
    Call UniCode_Conv(Y_GLICSREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '資材仕入先ﾜｰｸｾﾝﾀｰ
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN, KANKYO_KBN)                    '環境種類区分
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_ST, KANKYO_KBN_ST)              '環境種類区分適用開始
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)        '環境種類区分数量
    Call UniCode_Conv(Y_GLICSREC.ID_NO2, ID_NO2)                            'ID_NO
    Call UniCode_Conv(Y_GLICSREC.AITESAKI_CODE, AITESAKI_CODE)              '相手先ｺｰﾄﾞ
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_YMD, JYUCHU_YMD)                    '受注年月日
    Call UniCode_Conv(Y_GLICSREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)        '指定納期年月日
    Call UniCode_Conv(Y_GLICSREC.LIST_OUT_END_F, "")                        '入庫ﾘｽﾄ出力F
    Call UniCode_Conv(Y_GLICSREC.NYUKO_TANABAN, "")                         '入庫棚番
    Call UniCode_Conv(Y_GLICSREC.MAEGARI_SURYO, "")                         '前借相殺数
    '----------------   2010.07.08 △
    
    
    
    '2011.03.23 発生元プログラム
    Call UniCode_Conv(Y_GLICSREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
    '2011.03.23 元テキスト№
    If Trim(MOTO_TEXT_NO) = "" Then
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, "")
    Else
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
    End If
    
    Call UniCode_Conv(Y_GLICSREC.FILLER, "")

    Loop_Cnt = 0

    Do
        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)
            
            Case Else
                Call File_Error(sts, BtOpInsert, "入荷予定", 0)
                Exit Function
        End Select
    Loop

    Y_GLICS_PUT_PROC = False

End Function

'----------------------------------------------------------------------------
'           照合用入荷予定ファイル出力処理
'           2007.06.15
'----------------------------------------------------------------------------
'Dim sts     As Integer
'Dim ans     As Integer
'
'    Y_GLICS_PUT_PROC = True
'
'    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
'    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
'    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
'    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
'    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'
'
'    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
'    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
'    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
'
'    Call UniCode_Conv(Y_GLICSREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
'    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
'    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
'    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
'    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
'    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
'    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
'    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
'    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
'    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
'    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
'    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
'    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
'    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
'    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
'    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
'
'
'    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))
'
'
'                        '先行入荷数（入荷実績数）
'    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")
'
'                        '予算単位元
'    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'                        '予算単位先
'    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'                        '標準棚番
'    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
'    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'                        'H倉庫 2006.10.17
'    Call UniCode_Conv(Y_GLICSREC.H_SOKO, StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode))

'                        '入荷リスト出力フラグ   2007.06.12
'    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
'                        '直送区分
'    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, StrConv(New_HS_IN_SIJREC.CYOK_KBN, vbUnicode))
'                        '入出庫区分
'    Call UniCode_Conv(Y_GLICSREC.IO_KBN, StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode))
'                        '赤黒区分
'    Call UniCode_Conv(Y_GLICSREC.PM_KBN, StrConv(New_HS_IN_SIJREC.PM_KBN, vbUnicode))
'                        '伝票種別
'    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, StrConv(New_HS_IN_SIJREC.DEN_SYU, vbUnicode))
'                        '支給先／出荷先
'    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, StrConv(New_HS_IN_SIJREC.SYUK_CODE, vbUnicode))
'                        '支給先／出荷先名
'    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, StrConv(New_HS_IN_SIJREC.SYUK_NAME, vbUnicode))
'                        '挿入年月日
'    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
'
'
'    Call UniCode_Conv(Y_GLICSREC.FILLER, "")
'
'    Do
'        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'        Select Case sts
'            Case BtNoErr
'                Exit Do
'            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
'            Case Else
'                Call File_Error(sts, BtOpInsert, "入荷予定")
'                Exit Function
'        End Select
'    Loop
'
'    Y_GLICS_PUT_PROC = False
'
'
'End Function


