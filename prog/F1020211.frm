VERSION 5.00
Begin VB.Form F1020211 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入出荷予定データ取込み "
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
Attribute VB_Name = "F1020211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private WS_NO       As String * 2           'ﾜｰｸｽﾃｰｼｮﾝ番号


Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '仕向け先
    JGYOBU                  As String * 1   '事業部
    NAIGAI                  As String * 1   '国内外
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean

Private In_Cnt              As Integer      'データ読み込み件数
Private Out_Cnt             As Integer      'データ出力件数


Private Const In_Mode% = 1                  '入荷処理
Private Const Out_Mode% = 2                 '出荷処理

Private KASO_NYUKA_SOKO     As String * 2   '仮想入荷倉庫番号

Private GOODS_KBN           As String * 1   '商品化 要／不要


'''Private INS_DATE            As String * 8   '実行日付
'''Private INS_BIN             As Integer      '便

'Private Const LAST_UPDATE_DAY$ = "[F102021]2018.10.23 15:45 積水対応"
Private Const LAST_UPDATE_DAY$ = "[F102021]2019.01.16 13:30 積水対応"
                                            
                                        
Private Function Nyuka_Update_Proc(JGYOBU As String, ix As Integer) As Boolean
'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理
'----------------------------------------------------------------------------
Dim WK_Y_QTY        As Long     '出荷数ワーク
Dim WK_Qty          As Long     '前借残ワーク
Dim WK_E_QTY        As Long     '先行出荷数ワーク

Dim SUMI_QTY        As Long     '商品化済みとして登録
Dim MI_QTY          As Long     '未商品として登録

Dim DEN_NO          As String
Dim HINBAN          As String
Dim SURYO           As String
Dim ORDER_NO        As String
Dim ORDER_NO_1      As String
Dim ORDER_NO_2      As String

Dim ORDER_NO_T      As String
Dim ORDER_NO_I      As Integer

Dim NYUKO_YMD       As String
Dim NYUKO_YMD_T     As String


Dim sts             As Integer
Dim ans             As Integer
Dim Ret             As Integer
    
Dim HS_NYUKANo      As Long
Dim HS_NYUKA_OP     As Boolean
    
Dim FileName        As String
Dim Input_Wk        As Variant
Dim Input_Buffer    As String

Dim SKIP_F          As Boolean
Dim FAST_F          As Boolean
Dim NEXT_F          As Boolean

Dim c               As String * 128


Dim i               As Integer
    
    
    Nyuka_Update_Proc = True



    '入荷予定ファイル名取り込み & ＯＰＥＮ
    If GetIni("FILE", "HS_NYUKA", "SYS", c) Then
        Beep
        MsgBox "入荷予定ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)

    HS_NYUKA_OP = False

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_NYUKANo = FreeFile
    Open FileName For Input As #HS_NYUKANo
    On Error GoTo Exit_Proc
    HS_NYUKA_OP = True


    'ユニーク項目のﾀｲﾄﾙ取込み
    ORDER_NO_T = "注文№"
    If GetIni(App.EXEName, "ORDER_NO_T", App.EXEName, c) Then
    Else
        ORDER_NO_T = Trim(c)
    End If
    '入庫日のﾀｲﾄﾙ取込み
    NYUKO_YMD_T = "入庫日 : "
    If GetIni(App.EXEName, "NYUKO_YMD_T", App.EXEName, c) Then
    Else
        NYUKO_YMD_T = Trim(c)
    End If


    
    FAST_F = True
    NEXT_F = False


    Do While Not EOF(HS_NYUKANo)
        Line Input #HS_NYUKANo, Input_Buffer
        
        Input_Wk = Split(Input_Buffer, vbTab, -1)
    
    
    
        In_Cnt = In_Cnt + 1
        lblINCNT(ix).Caption = Format(In_Cnt, "#0")
        DoEvents
    
        SKIP_F = False
    
    
        If FAST_F Then
            
            
            If UBound(Input_Wk) > 4 Then
            
                If InStr(1, Input_Wk(5), NYUKO_YMD_T) > 0 Then
            
            
                    NYUKO_YMD = Format(Right(Input_Wk(5), 11), "YYYYMMDD")
            
                End If
            
            End If
            
            
            
            
            For i = 0 To UBound(Input_Wk)
                If Trim(Input_Wk(i)) = ORDER_NO_T Then
                    ORDER_NO_I = i
                    FAST_F = False
                    Exit For
                End If
            Next i
        End If
    
        If Not FAST_F Then
            If UBound(Input_Wk) < 4 Then
                SKIP_F = True
            Else
                If NEXT_F Then
                    ORDER_NO_2 = Trim(Input_Wk(4))
                    NEXT_F = False
                Else
                    If Not IsNumeric(Trim(Input_Wk(3))) Then
                        SKIP_F = True
                    Else
                        DEN_NO = Trim(Input_Wk(0))
                        HINBAN = Trim(Input_Wk(2))
                        SURYO = Trim(Input_Wk(3))
                        ORDER_NO_1 = Trim(Input_Wk(4))
                
                        NEXT_F = True
                    End If
                End If
            End If
        End If
        
            
        
        
                
        
        
        
        
        
        If Not SKIP_F And Not FAST_F And Not NEXT_F Then
                                            
            ORDER_NO = Trim(ORDER_NO_1) & Trim(ORDER_NO_2)
                                            
            '入荷予定重複チェック
            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, NYUKO_YMD)
            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, ORDER_NO)
            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & ORDER_NO)
'''                    SKIP_F = True
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "入荷予定")
                    Exit Function
            End Select
                                            
                                            
            If Not SKIP_F Then
                                                'トランザクション開始
                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                    Exit Function
                End If
                                            '品目マスタチェック
                If Item_Check_Proc(In_Mode, JGYOBU, NAIGAI_NAI, HINBAN) Then
                    GoTo Abort_Tran
                End If
                                            
                                            '入荷データ作成
                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, ORDER_NO)
        
        
                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NO, HINBAN)
                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(SURYO), "0000000"))
                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, NYUKO_YMD)
                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, NYUKO_YMD)
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
                Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
        
        
        
        
                '入荷ﾁｪｯｸﾃﾞｰﾀ更新
                Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
                Call UniCode_Conv(K0_J_NYU.NAIGAI, NAIGAI_NAI)
                Call UniCode_Conv(K0_J_NYU.HIN_GAI, HINBAN)
    
                WK_Y_QTY = CLng(SURYO)
    
    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                        
                                Do
                                
                                    sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                            Beep
                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                            If ans = vbCancel Then
                                                Exit Function
                                            End If
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                                            Exit Function
                                    End Select
                                
                                Loop
                                WK_E_QTY = CLng(SURYO)
                            Else
                                Do
                                    sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                            Beep
                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                            If ans = vbCancel Then
                                                Exit Function
                                            End If
                                        Case Else
                                            Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ")
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
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                           End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                            Exit Function
                    End Select
                Loop
                                    '先行入荷数（入荷実績数）
                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
        
                                    '予算単位元
                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, "")
                                    '予算単位先
                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, "")
                                    '標準棚番
                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_NAI, "")
                                    'H倉庫 2006.10.17
                Call UniCode_Conv(Y_NYUREC.H_SOKO, "")
    
                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                
                Do
                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "入荷予定")
                            Exit Function
                    End Select
                Loop
        
                        
                
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                    MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                    SUMI_QTY = 0
                Else
                    SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                    MI_QTY = 0
                End If
                
        
                '入荷数で在庫データ更新（＋）
                If Nyuko_Update_Proc(JGYOBU, _
                                    NAIGAI_NAI, _
                                    HINBAN, _
                                    NYUKO_YMD, _
                                    (KASO_NYUKA_SOKO & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, MI_QTY, _
                                    WS_NO, WS_NO, , _
                                    NYUKO_YMD & " 伝№:" & DEN_NO) Then
                    Exit Function
            
                End If
            
                '前借り数で在庫データ更新（－）
                If WK_E_QTY <> 0 Then
                '在庫データLOCK
                    If Zaiko_Lock_Proc((KASO_NYUKA_SOKO & "01" & "01" & "01"), _
                                        JGYOBU, _
                                        NAIGAI_NAI, _
                                        HINBAN, _
                                        WS_NO) Then
                        Exit Function
    
                    End If
        
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                        MI_QTY = WK_E_QTY
                    Else
                        SUMI_QTY = WK_E_QTY
                    End If
            
            
                    If Syuko_Update_Proc(JGYOBU, _
                                        NAIGAI_NAI, _
                                        HINBAN, _
                                        NYUKO_YMD, _
                                        (KASO_NYUKA_SOKO & "01" & "01" & "01"), _
                                        YOIN_MAE_SOUSAI, _
                                        SUMI_QTY, MI_QTY, 0, _
                                        WS_NO, WS_NO) Then
                        Exit Function
        
                    End If
                End If
        
                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    GoTo Abort_Tran
                End If

        
        
        
            End If
            
            
            
            
            Out_Cnt = Out_Cnt + 1
            lblOUTCNT(ix).Caption = Format(Out_Cnt, "#0")
            DoEvents

        End If
    
    Loop

    Nyuka_Update_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


Exit_Proc:
    
    If HS_NYUKA_OP Then
        Close #HS_NYUKANo
    End If
    

End Function
    
Private Function Syuka_Update_Proc(JGYOBU As String, ix As Integer) As Boolean
'----------------------------------------------------------------------------
'                   「出荷予定データ」更新処理
'----------------------------------------------------------------------------



Dim INS_NOW         As String


Dim sts             As Integer
Dim Ret             As String

Dim DUP_SYUKANo     As Long

Dim HS_SMEISAINo    As Long
Dim HS_SMEISAI_OP   As Boolean

Dim HS_PICNo        As Long
Dim HS_PIC_OP       As Boolean

Dim FileName        As String

Dim c               As String * 128

Dim i               As Integer

Dim Input_Buffer    As String
Dim Pos             As Integer
        
Dim Skip_Flg        As Boolean
Dim Fast_Flg        As Boolean

Dim Input_Wk        As Variant

Dim Location        As String
Dim HIN_NAME        As String

Dim SYUKA_NO        As String
Dim SYUKA_YMD       As String


Dim COL_OKURISAKI_CD _
                    As String
Dim OKURISAKI_CD    As String


Dim wkOKURISAKI_CD  As String
Dim svOKURISAKI_CD  As String

Dim OKURISAKI       As String

Dim svOKURISAKI     As String


Dim URIDEN          As String
Dim DEN_NO          As String
Dim HINBAN          As String
Dim SURYO           As String
Dim CYU_NO          As String
Dim TOKUI_CODE      As String
Dim TOKUI_NAME      As String
Dim BIKOU           As String
Dim UNSOU           As String
Dim INS_BIN         As String               '2007.01.16
Dim JYUSHO          As String               '2009.11.19


Dim YUBIN_NO        As String               '2010.04.05
Dim TEL_NO          As String               '2010.04.05


Dim SEK_KEN_NO      As String               '件管№　　　■管理№(上)   2011.04.30
Dim SEK_HIN_NO      As String               '品管№　　　■管理№(下)   2011.04.30



Dim SV_DEN_NO       As String * 7


Dim DEN_SEQ         As Integer

Dim ID_SET_FLG      As Boolean
Dim SV_ID_NO        As String * 7
Dim ID_SEQ          As Integer



Dim ans             As Integer

Dim DUP_Flg         As Boolean              '2007.07.31



    Syuka_Update_Proc = False

    '出荷明細ファイル名取り込み & ＯＰＥＮ
    If GetIni("FILE", "HS_SMEISAI", "SYS", c) Then
        Beep
        MsgBox "出荷明細ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)

    HS_SMEISAI_OP = False

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_SMEISAINo = FreeFile
    Open FileName For Input As #HS_SMEISAINo

    On Error GoTo Exit_Proc            '処理終了
    HS_SMEISAI_OP = True
    
    'ピッキングファイル名取り込み & ＯＰＥＮ
    If GetIni("FILE", "HS_PIC", "SYS", c) Then
        Beep
        MsgBox "ピッキングリスト・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)

    HS_PIC_OP = False
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_PICNo = FreeFile
    Open FileName For Input As #HS_PICNo

    On Error GoTo Exit_Proc            '処理終了
    HS_PIC_OP = True
    
    Syuka_Update_Proc = True
    
    
    '出荷重複ファイル名取り込み & ＯＰＥＮ
    
    If GetIni("FILE", "SYUDUP  ", "SYS", c) Then
        Beep
        MsgBox "出荷重複ファイル・ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Exit Function
    End If
    FileName = Trim(c)
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)




    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")



    '-----------------------------------------------------------------  ピッキングリストより品目マスタ作成／更新
    Do While Not EOF(HS_PICNo)
        
        DoEvents
        
        Line Input #HS_PICNo, Input_Buffer
        
        Input_Wk = Split(Input_Buffer, vbTab, -1)
            
        Location = ""
        HINBAN = ""
        HIN_NAME = ""
    
    
        If UBound(Input_Wk) > 6 Then
            Location = StrConv(Input_Wk(1), vbNarrow)
            HINBAN = Input_Wk(3)
            HIN_NAME = Input_Wk(7)
        End If
    
        If Trim(HINBAN) = "" Or _
            Trim(HIN_NAME) = "" Then
        Else
                        '品目マスタチェック
            If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HINBAN, HIN_NAME, Location) Then
                Exit Function
            End If
        
        
        End If
    Loop




    svOKURISAKI_CD = ""
    Fast_Flg = True

    DUP_Flg = False


    Do While Not EOF(HS_SMEISAINo)
        
        In_Cnt = In_Cnt + 1
        lblINCNT(ix).Caption = Format(In_Cnt, "#0")
        DoEvents
        
        
        
        Line Input #HS_SMEISAINo, Input_Buffer




        Input_Wk = Split(Input_Buffer, vbTab, -1)
Debug.Print "Input_Wk(1)=" & Input_Wk(1)

        SYUKA_NO = ""
        SYUKA_YMD = ""
        
        
        COL_OKURISAKI_CD = ""
        OKURISAKI_CD = ""

        
        OKURISAKI = ""
        URIDEN = ""
        DEN_NO = ""
        HINBAN = ""
        SURYO = ""
        CYU_NO = ""
        TOKUI_CODE = ""
        TOKUI_NAME = ""
        BIKOU = ""
        UNSOU = ""
        INS_BIN = ""
        JYUSHO = ""
        
        
        YUBIN_NO = ""
        TEL_NO = ""
        
        
        SEK_KEN_NO = ""         '件管№　　　■管理№(上)   2011.04.30
        SEK_HIN_NO = ""         '品管№　　　■管理№(下)   2011.04.30
        
        
        
        '出荷№
        If UBound(Input_Wk) > 0 - 1 Then
            SYUKA_NO = Input_Wk(1 - 1)
        End If
        
        
        If Not IsNumeric(SYUKA_NO) Then
        Else
            '出荷日
            If UBound(Input_Wk) > 1 - 1 Then
                
'>>>>>>>>>>>>>>>>>>>>>>     2018.10.23
                If CStr(Input_Wk(2 - 1)) > Mid(Format(Now, "YYYY/MM/DD"), 6, 5) Then
                    SYUKA_YMD = Format(CLng(Mid(Format(Now, "YYYYMMDD"), 1, 4) - 1), "0000") & "/" & Input_Wk(2 - 1)
                Else
'>>>>>>>>>>>>>>>>>>>>>>     2018.10.23
                
                
                
                
                
                
                    If Mid(Format(Now, "YYYYMMDD"), 5, 2) = "12" Then
                        If Mid(CStr(Input_Wk(2 - 1)), 1, 2) = "01" Then
                            SYUKA_YMD = Format(CLng(Mid(Format(Now, "YYYYMMDD"), 1, 4) + 1), "0000") & "/" & Input_Wk(2 - 1)
                        Else
                            SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2 - 1)
                        End If
                    Else
                        SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2 - 1)
                    End If
                End If      '>>>>>>>>>>>>>>>>>>>>>>     2018.10.23
            End If
        
        
If Trim(Input_Wk(5)) = "Z0000003" Then
    Debug.Print
End If
        
        
            '集約送り先CD
            If UBound(Input_Wk) > 3 - 1 Then
            
                wkOKURISAKI_CD = Trim(Input_Wk(4 - 1))
            
                COL_OKURISAKI_CD = Trim(Input_Wk(4 - 1))
            
                If UBound(Input_Wk) > 4 - 1 Then
                    OKURISAKI_CD = Trim(Input_Wk(5 - 1))
                    If Trim(wkOKURISAKI_CD) = "" Then
                        wkOKURISAKI_CD = Trim(Input_Wk(5 - 1))
                    
                    End If
            
            
                    ID_SET_FLG = False
            
            
                    If Trim(wkOKURISAKI_CD) <> Trim(svOKURISAKI_CD) Then

                        svOKURISAKI_CD = wkOKURISAKI_CD
                        ID_SET_FLG = True

                    End If
            
                End If
            
            End If
        
            '送り先名
            If UBound(Input_Wk) > 6 - 1 Then
                If Trim(Input_Wk(7 - 1)) = "" Then
                Else
                    OKURISAKI = Trim(Input_Wk(7 - 1))
                End If
            
            
                If Trim(OKURISAKI) <> "" Then
                    svOKURISAKI = OKURISAKI
                End If
            End If
            
            
            '売伝
            If UBound(Input_Wk) > 7 - 1 Then
                URIDEN = Input_Wk(8 - 1)
            End If
            '伝票番号
            If UBound(Input_Wk) > 9 - 1 Then
                
                If Len(Input_Wk(10 - 1)) > 7 - 1 Then

                    DEN_NO = Left(Input_Wk(10 - 1), 7)
                Else
                    DEN_NO = Trim(Input_Wk(10 - 1))
                End If
            
                If ID_SET_FLG Then
                    SV_ID_NO = DEN_NO
                    ID_SEQ = 0
                End If
            
            End If
            '品番
            If UBound(Input_Wk) > 11 - 1 Then
                HINBAN = Trim(Input_Wk(12 - 1))
            End If
            '数量
            If UBound(Input_Wk) > 12 - 1 Then
                SURYO = Trim(Input_Wk(13 - 1))
            End If
            '注文№
            If UBound(Input_Wk) > 14 - 1 Then
                CYU_NO = Trim(Input_Wk(15 - 1))
            End If
            '得意先ｺｰﾄﾞ
            If UBound(Input_Wk) > 15 - 1 Then
                TOKUI_CODE = Trim(Input_Wk(16 - 1))
            End If
            '得意先名
            If UBound(Input_Wk) > 16 - 1 Then
                TOKUI_NAME = Trim(Input_Wk(17 - 1))
            End If
            '備考
            If UBound(Input_Wk) > 18 - 1 Then
                BIKOU = Trim(Input_Wk(19 - 1))
            End If
            '運送会社
            If UBound(Input_Wk) > 20 - 1 Then
                UNSOU = Trim(Input_Wk(21 - 1))
            End If
            '便 '2007.01.16
            If UBound(Input_Wk) > 21 - 1 Then
                INS_BIN = Trim(Input_Wk(22 - 1))
            End If
            
            
            '住所 '2009.11.19
            If UBound(Input_Wk) > 22 - 1 Then
                JYUSHO = Trim(Input_Wk(23 - 1))
            End If
            
            
            '郵便番号 '2010.04.05
            If UBound(Input_Wk) > 23 - 1 Then
                YUBIN_NO = Trim(Input_Wk(24 - 1))
            End If
            
            
            '電話番号 '2010.04.05
            If UBound(Input_Wk) > 24 - 1 Then
                TEL_NO = Trim(Input_Wk(25 - 1))
            End If
            
            
            '件管№　　　■管理№(上)   2011.04.30
            If UBound(Input_Wk) > 25 - 1 Then
                SEK_KEN_NO = Trim(Input_Wk(26 - 1))
            End If
            
            '品管№　　　■管理№(下)   2011.04.30
            If UBound(Input_Wk) > 26 - 1 Then
                SEK_HIN_NO = Trim(Input_Wk(27 - 1))
            End If
            
            'ｴﾗｰﾁｪｯｸ
            Skip_Flg = False
            
            If Trim(SYUKA_YMD) = "" Or _
                Trim(DEN_NO) = "" Or _
                Trim(HINBAN) = "" Or _
                Trim(SURYO) = "" Then
                
                Skip_Flg = True
        
            Else
        
                If Not IsDate(SYUKA_YMD) Then
                    Skip_Flg = True
                Else
                    SYUKA_YMD = (Format(SYUKA_YMD, "YYYYMMDD"))
                End If
        
                If Not IsNumeric(SURYO) Then
                    Skip_Flg = True
                Else
                    If CLng(SURYO) = 0 Then
                        Skip_Flg = True
                    End If
                End If
        
        
            End If
        
            If Not Skip_Flg Then
                
                
                If Trim(SV_DEN_NO) = "" Then
                    SV_DEN_NO = DEN_NO
                    DEN_SEQ = 0
                End If
        
                If SV_DEN_NO <> DEN_NO Then
                    SV_DEN_NO = DEN_NO
                    DEN_SEQ = 0
                End If
                
                DEN_SEQ = DEN_SEQ + 1
                ID_SEQ = ID_SEQ + 1
        
            
        
        
                '出荷予定重複ﾁｪｯｸ
                Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
        
                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Skip_Flg = True
                
                
                        If Fast_Flg Then
                            DUP_SYUKANo = FreeFile
                            Open FileName For Append As #DUP_SYUKANo

                            Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")
                            Write #DUP_SYUKANo, "№", "出荷日", "送り先名", "売伝", "伝票番号", "品番", "数量", "注文№", "得意先CD", "得意先名", "備考", "運送会社"
                            Fast_Flg = False
                        
                        End If
                
                
                        Write #DUP_SYUKANo, SYUKA_NO,
                        Write #DUP_SYUKANo, SYUKA_YMD,
                        Write #DUP_SYUKANo, OKURISAKI,
                        Write #DUP_SYUKANo, URIDEN,
                        Write #DUP_SYUKANo, DEN_NO,
                        Write #DUP_SYUKANo, HINBAN,
                        Write #DUP_SYUKANo, SURYO,
                        Write #DUP_SYUKANo, CYU_NO,
                        Write #DUP_SYUKANo, TOKUI_CODE,
                        Write #DUP_SYUKANo, TOKUI_NAME,
                        Write #DUP_SYUKANo, BIKOU,
                        Write #DUP_SYUKANo, UNSOU,
                        Write #DUP_SYUKANo, JYUSHO



                
                        Call LOG_OUT(LOG_F, "Y_SYUKA.DAT DUP 事業部=" & JGYOBU & " 伝票ＩＤ＝" & DEN_NO & " 品番=" & HINBAN)
                
                        DUP_Flg = True
                
                        GoTo NEXT_LOOP
                        
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定")
                        Exit Function
                End Select
        
        
                If Not Skip_Flg Then
                
                
                    'ﾄﾗﾝｻﾞｸｼｮﾝ開始
                    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                        Exit Function
                    End If
                    '---------------------------------------------------------- 得意先のﾁｪｯｸ
                    Call UniCode_Conv(K0_MTS.MUKE_CODE, TOKUI_CODE)
                    Call UniCode_Conv(K0_MTS.SS_CODE, "")
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            '未登録は自動作成
                            Call UniCode_Conv(MTSREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(MTSREC.DATA_KBN, "")
                            Call UniCode_Conv(MTSREC.MUKE_CODE, TOKUI_CODE)
                            Call UniCode_Conv(MTSREC.SS_CODE, "")
                            Call UniCode_Conv(MTSREC.MUKE_NAME, TOKUI_NAME)
                            Call UniCode_Conv(MTSREC.SS_NAME, "")
                            Call UniCode_Conv(MTSREC.MUKE_DNAME, TOKUI_NAME)
                            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "")
                            Call UniCode_Conv(MTSREC.FILLER, "")
                            Do
                                sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                        Beep
                                        ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                        If ans = vbCancel Then
                                            GoTo Abort_Tran
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpInsert, "向け先管理ﾏｽﾀ" & "key=" & TOKUI_CODE)
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                                        
                                                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            GoTo Abort_Tran
                    End Select
                
                    '---------------------------------------------------------- 品目ﾏｽﾀのﾁｪｯｸ
                    If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HINBAN) Then
                        GoTo Abort_Tran
                    End If
                
                
                    '---------------------------------------------------------- 出荷予定作成
                
                
                    '使用端末ID
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    '使用中ﾌﾟﾛｸﾞﾗﾑID
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    '完了区分
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    'ﾃﾞｰﾀ種別
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    '事業部
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    '注文区分(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                    'ID-NO
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
                    '国内外
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                    '品番(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HINBAN)
                    '得意先(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, TOKUI_CODE)
                    '直送先(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                    '出荷日(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
                    '事業場ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                    'ﾃﾞｰﾀ区分
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                    '取引区分
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                    'ID-NO
                    Call UniCode_Conv(Y_SYUREC.ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
                    '会計用事業場ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                    '品番
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, HINBAN)
                    '伝票番号
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, SV_DEN_NO)
                    '出荷数量
                    Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(SURYO), "0000000"))
                    '得意先
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, TOKUI_CODE)
                    '出庫収支
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")
                    '資産管理用在庫収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                    '補助在庫収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                    '出荷日付
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
                    '実際単価
                    Call UniCode_Conv(Y_SYUREC.TANKA, "")
                    'ｵｰﾀﾞｰ番号
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                    'ｱｲﾃﾑ番号
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                    '注文管理番号略号
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                    '個装形態ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                    '出荷予定日
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
                    'ﾛｹｰｼｮﾝ1
                    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                    'ﾛｹｰｼｮﾝ2
                    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                    'ﾛｹｰｼｮﾝ3
                    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    '得意先名称
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, TOKUI_NAME)
                    '注文区分
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                    '注文区分名称
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_2)
                    '原産国1
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                    '原産国2
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                    '備考2
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                    '販売区分
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                    '直送指示区分
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")
                    'ﾕﾆｯﾄ修正管理番号
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                    '在庫引当順序
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                    '合梱管理番号
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                    '受注残数量
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                    '供給区分
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                    '商品化納品在庫収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                    '商品化納品資産管理収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                    '商品化納品補助収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                    '備考1
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                    '帳端区分
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                    '受付品目番号
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                    '品名
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    '品目番号変更区分
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                    'ﾓｼﾞｭｰﾙ交換区分
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                    '残在庫まとめ在庫収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                    '残在庫まとめ資産管理収支ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                    '指定納期
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                    'ｻｰﾋﾞｽ会社管理番号
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                    '機種品目ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                    '環境企画部品区分
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                    '直送相手先ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                    '欠品解消区分
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
                    '品番（内部）
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    '標準棚番
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
                    '出庫表印刷日付
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                    '完了日付
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    '検品日付
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    '特売り区分
                    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                    '出庫実績数量
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")
                    '取込み日時
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    '検品担当者ｺｰﾄﾞ
                    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")
                    '検品時刻
                    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")
                    '上位ﾘﾝｸ用向け先
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, TOKUI_CODE)
                    '上位ﾘﾝｸ用連番
                    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")
                    '画面検品ﾌﾗｸﾞ
                    Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")
                    '検品時数量
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")
                     '入出庫区分(引数、処理区分) 2017.10.19
                    Call UniCode_Conv(Y_SYUREC.H_IO_KBN, "")
                    '倉庫ｺｰﾄﾞ(在庫収支ﾃｰﾌﾞﾙの倉庫ｺｰﾄﾞ) 2017.10.19
                    Call UniCode_Conv(Y_SYUREC.H_SOKO_CODE, "")
                    '更新日時   2017.10.19
                    Call UniCode_Conv(Y_SYUREC.UPD_NOW, "")
                     '完了時刻   2017.10.19
                    Call UniCode_Conv(Y_SYUREC.KAN_HMS, "")
                    
                    '
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
                                     GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "出荷予定")
                                GoTo Abort_Tran
                        End Select
                    Loop
    
    
    
                    '---------------------------------------------------------- 出荷予定(ﾎｽﾄｲﾒｰｼﾞ)作成
                    'ID-NO
                    Call UniCode_Conv(Y_SYU_HREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
                    '№
                    Call UniCode_Conv(Y_SYU_HREC.SYUKA_NO, SYUKA_NO)
                    '出荷日付
                    Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, SYUKA_YMD)
                    '送り先名
                    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, svOKURISAKI)
                    '売伝
                    If Trim(URIDEN) = "" Then
                        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "0")
                    Else
                        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "1")
                    End If
                    '伝票番号
                    Call UniCode_Conv(Y_SYU_HREC.DEN_NO, SV_DEN_NO)
                    '追番
                    Call UniCode_Conv(Y_SYU_HREC.SEQ_NO, Format(DEN_SEQ, "0"))
                    '品番
                    Call UniCode_Conv(Y_SYU_HREC.HIN_NO, HINBAN)
                    '数量
                    Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(SURYO), "0000000"))
                    '注文№
                    Call UniCode_Conv(Y_SYU_HREC.ODER_NO, CYU_NO)
                    '得意先
                    Call UniCode_Conv(Y_SYU_HREC.MUKE_CODE, TOKUI_CODE)
                    '得意先名称
                    Call UniCode_Conv(Y_SYU_HREC.MUKE_NAME, TOKUI_NAME)
                    '備考
                    Call UniCode_Conv(Y_SYU_HREC.BIKOU, BIKOU)
                    '運送会社名
                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU)
                    '取込み日時
                    Call UniCode_Conv(Y_SYU_HREC.INS_NOW, INS_NOW)
                    '出荷ﾗﾍﾞﾙ印刷日時
                    Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, "")
                    'ﾃﾞｰﾀ発生順
                    Call UniCode_Conv(Y_SYU_HREC.DATA_CNT, Format(Out_Cnt, "00000"))
                    '送り状№
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
                    '検品日時
                    Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
                    '検品担当者
                    Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
                    '口数
                    Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "0000")   '2007.02.01
                    Call UniCode_Conv(Y_SYU_HREC.xKUTI_SU, "00")    '2007.02.01
                    
                    '強制完了
                    Call UniCode_Conv(Y_SYU_HREC.KYOSEI_END, "")
                    'ｷｬﾝｾﾙF
                    Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "")
                    '備考
                    Call UniCode_Conv(Y_SYU_HREC.INPUT_BIKOU, "")
                    '便 2007.01.16
                    If IsNumeric(INS_BIN) Then
                        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, Format(CInt(INS_BIN), "00"))
                    Else
                        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, "")
                    End If
                    
                    
                    '事業部 2007.03.14
                    Call UniCode_Conv(Y_SYU_HREC.JGYOBU, JGYOBU)
                    '国内外 2007.03.14
                    Call UniCode_Conv(Y_SYU_HREC.NAIGAI, NAIGAI_NAI)
                    '出庫表№ 2007.03.14
                    Call UniCode_Conv(Y_SYU_HREC.SYU_NO, "")
                    '出庫実績数 2007.03.14
                    Call UniCode_Conv(Y_SYU_HREC.J_SURYO, "0000000")
                                        
                    '集約送り先CD
                    Call UniCode_Conv(Y_SYU_HREC.COL_OKURISAKI_CD, COL_OKURISAKI_CD)
                    '送り先CD
                    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI_CD, OKURISAKI_CD)
                    
                    
                    '住所 2009.11.19
                    Call UniCode_Conv(Y_SYU_HREC.JYUSHO, JYUSHO)
                    
                    
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.TEL_NO, TEL_NO)            '電話番号   2010.01.21
                    Call UniCode_Conv(Y_SYU_HREC.YUBIN_NO, YUBIN_NO)        '郵便番号   2010.01.21
                    Call UniCode_Conv(Y_SYU_HREC.JURYO, "")                     '重量   　　2010.01.21
                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU, "")                    '才数   　　2010.01.21
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO_SEQ, "")              '送り状№　枝番　2010.01.21
    
                    '梱包Ｆ 2010.02.13
                    Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, StrConv(ITEMREC.KONPOU_F, vbUnicode))
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, "")              '口数(単体)     2010.01.21
                    
                    Call UniCode_Conv(Y_SYU_HREC.SEK_KEN_NO, SEK_KEN_NO)       '件管№　　　■管理№(上)   2011.04.30
                    Call UniCode_Conv(Y_SYU_HREC.SEK_HIN_NO, SEK_HIN_NO)       '品管№　　　■管理№(下)   2011.04.30
                    
                    
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_TANTO, "")           '注文ﾃﾞｰﾀ照合担当       2011.05.02
                    Call UniCode_Conv(Y_SYU_HREC.SEK_SHOGO_DATETIME, "")        '注文ﾃﾞｰﾀ照合日時       2011.05.02
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.CNT_BARA_SU, "")               '検品実績　バラ     2012.10.02
                    Call UniCode_Conv(Y_SYU_HREC.CNT_HAKO_SU, "")               '検品実績　箱       2012.10.02
                    Call UniCode_Conv(Y_SYU_HREC.GAISO_IRI_QTY, "")             '外装入り数         2012.10.02
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.Y_HIN_CHK_CNT, "")             '品番読込み回数     2012.10.02
                    Call UniCode_Conv(Y_SYU_HREC.J_HIN_CHK_CNT, "")             '品番読込み済み回数 2012.10.02
                    
                    Call UniCode_Conv(Y_SYU_HREC.KEN_HINBAN, "")                '検品中品番         2012.10.02
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.TYAKUTEN, "")                  '着店コード         2017.02.08
                   
                    
                    
                    
                    
                    Call UniCode_Conv(Y_SYU_HREC.FILLER, "")
                    
                                                                                '追加担当者
                    Call UniCode_Conv(Y_SYU_HREC.INS_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                                '追加日時
                    Call UniCode_Conv(Y_SYU_HREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                                                                                '更新担当者
                    Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, "")
                                                                                '更新日時
                    Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, "")
                    
                    
                    
                    
                    
                    
                    
                    Do
                        sts = BTRV(BtOpInsert, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                     GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                                GoTo Abort_Tran
                        End Select
                    Loop
                    
                    
                    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
                
                
                
                
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(ix).Caption = Format(Out_Cnt, "#0")
                
                
                
                
                
                
                End If
        
        
            End If
        
        
        End If

NEXT_LOOP:


    Loop















    Close #HS_SMEISAINo
    Close #HS_PICNo
    
    
    
    If DUP_Flg Then
        MsgBox "重複データが有りました。ログ(SYUKA_DUPyymmdd)の内容を確認して下さい"
    End If
    
    
    
    Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    Exit Function
    
Exit_Proc:
    
    
''''MsgBox Err.Number
    
    
    If HS_SMEISAI_OP Then
        Close #HS_SMEISAINo
    End If
    
    If HS_PIC_OP Then
        Close #HS_PICNo
    End If
    
    
End Function

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String


    '---------------------------------------------  事業部毎メインループ
    
    In_Cnt = 0
    Out_Cnt = 0
    
    For i = 0 To UBound(JGYOBU_T)
        

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = "0"
        lblINCNT(i).Caption = "0"
        DoEvents

'''2007.06.06        If Nyuka_Update_Proc(JGYOBU_T(i).CODE, i) Then    '入荷予定データ更新処理
'''2007.06.06
'''2007.06.06            Unload Me
'''2007.06.06
'''2007.06.06        End If
    
    
    
    
    
        If Syuka_Update_Proc(JGYOBU_T(i).CODE, i) Then '出荷予定データ更新処理

            Unload Me
        End If
    
    
    
    
    
    
    Next i


    Unload Me

Error_Proc:

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

    F1020211.Caption = F1020211.Caption & LAST_UPDATE_DAY


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
                               
    If JGYOB_TB_Set(1) Then     '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
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
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If



'   ---入庫関係処理による追加 2007.03.19
                                '倉庫ﾏｽﾀＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者ﾏｽﾀＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因ﾏｽﾀＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番ﾏｽﾀＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入庫実績(前借)ＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '資材 入庫実績(前借)ＯＰＥＮ
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '作業ﾛｸﾞＯＰＥＮ
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If



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
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
                                '便の獲得       '2007.01.16
'''    If GetIni(App.EXEName, "INS_DATE", App.EXEName, c) Then
'''        INS_DATE = Format(Now, "YYYYMMDD")
'''        INS_BIN = 1
'''    Else
'''        If Trim(c) <> Format(Now, "YYYYMMDD") Then
'''            INS_DATE = Format(Now, "YYYYMMDD")
'''            INS_BIN = 1
'''        Else
'''            INS_DATE = Trim(c)
'''
'''            If GetIni(App.EXEName, "INS_BIN", App.EXEName, c) Then
'''                INS_BIN = 1
'''            Else
'''                If IsNumeric(Trim(c)) Then
'''                    INS_BIN = CInt(Trim(c)) + 1
'''                Else
'''                    INS_BIN = 1
'''                End If
'''            End If
'''        End If
'''    End If
'''
'''                                'ＩＮＩ 本日日付出力
'''    If WriteIni(App.EXEName, "INS_DATE", App.EXEName, INS_DATE) Then
'''    End If
'''                                'ＩＮＩ 便出力
'''    If WriteIni(App.EXEName, "INS_BIN", App.EXEName, Format(INS_BIN, "0")) Then
'''    End If
    
    
    
                                '商品化要／不要の獲得
    GOODS_KBN = "0"
    If GetIni(App.EXEName, "GOODS_KBN", App.EXEName, c) Then
    Else
        GOODS_KBN = Trim(c)
    End If
    
    
    


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り？
'        Call Last_Proc
'    End If

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
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            
'   ---入庫関係処理による追加 2007.03.19
                                '倉庫ﾏｽﾀＯＰＥＮ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫ﾏｽﾀ")
        End If
    End If
                                '担当者ﾏｽﾀＯＰＥＮ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者ﾏｽﾀ")
        End If
    End If
                                '要因ﾏｽﾀＯＰＥＮ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因ﾏｽﾀ")
        End If
    End If
                                '品目ﾏｽﾀＯＰＥＮ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                '入荷予定ＯＰＥＮ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定")
        End If
    End If
                                '在庫ﾃﾞｰﾀＯＰＥＮ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫ﾃﾞｰﾀ")
        End If
    End If
                                '在庫移動歴ＯＰＥＮ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                
                                
                                '入庫実績(前借)ＯＰＥＮ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入庫実績")
        End If
    End If
                                '資材 入庫実績(前借)ＯＰＥＮ
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材 入庫実績")
        End If
    End If
                                '作業ﾛｸﾞＯＰＥＮ
    sts = BTRV(BtOpClose, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "作業ﾛｸﾞ")
        End If
    End If
                                            
                                            
                                            
                                            
                                            
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020211 = Nothing

    End
End Sub

Private Function Item_Check_Proc(Mode As Integer, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional HIN_NAME As String = "", _
                                    Optional Location As String = "") As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ」チェック＆更新処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer
        
Dim i           As Integer
    
    
    Item_Check_Proc = True

           

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
                com = BtOpUpdate
                                
                If Trim(HIN_NAME) <> "" Then
                    Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)   '品名
                End If
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)           '事業部
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)           '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI)         '品番（外部）
    
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)       '品名
    
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")            '標準棚番設定日
                
                
                                                                    '標準棚番
                If Len(Trim(Location)) > 6 Then
                    Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(Location, 1, 2))
                    Call UniCode_Conv(ITEMREC.ST_RETU, Mid(Location, 3, 2))
                    Call UniCode_Conv(ITEMREC.ST_REN, Mid(Location, 5, 2))
                    Call UniCode_Conv(ITEMREC.ST_DAN, "01")
                
                Else
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                End If
    
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")             '前回入庫倉庫
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")          '最終入庫日
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")          '最終出庫日
    
                Call UniCode_Conv(ITEMREC.HIN_NAI, "")              '品番（内部）
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
                
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
                
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
                
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
                
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
                
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_KBN)     '商品化有無（有）
                
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
                
                Call UniCode_Conv(ITEMREC.RANK, "")                 '現在ﾗﾝｸ
                Call UniCode_Conv(ITEMREC.NEW_RANK, "")             '新ﾗﾝｸ
                
                
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                
                
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
                
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
                Call UniCode_Conv(ITEMREC.xL_KISHU2, "")                '           機種(2)未使用
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
                Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)        '在庫管理対象有無　（対象）
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '機種(2)
                
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000")  '前月在庫数
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "00000000") '最終出荷数
                            
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")     'S2 在庫
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")     'P2 在庫
                            
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
                            
                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給区分
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給区分
                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価
    
    
                Call UniCode_Conv(ITEMREC.KUTI_SU, "")                  '口数               2010.01.18
                Call UniCode_Conv(ITEMREC.KONPOU_F, "")                 '梱包区分           2010.01.18
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2010.01.18
    
    
    
    
                            
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '更新担当者
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '更新日時
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                Exit Function
        End Select
    Loop
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "品目マスタ")
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
                
                
                
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "構成マスタ")
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If

    Item_Check_Proc = False

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

