VERSION 5.00
Begin VB.Form F1200121 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ホスト棚番設定データ作成"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2250
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   18
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1800
      Width           =   615
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   39
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   38
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   37
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   36
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   35
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   34
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   33
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   32
      Top             =   4440
      Width           =   915
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   31
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   30
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   29
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   28
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   27
      Top             =   4440
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   26
      Top             =   3960
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   25
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   24
      Top             =   3000
      Width           =   915
   End
   Begin VB.Label lblOUT_CNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   23
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label lblin_Cnt 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   22
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "出力件数＝"
      Height          =   255
      Index           =   5
      Left            =   2400
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(指定した日付以降の品目を出力します。)"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   20
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   19
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   1920
      Width           =   255
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
      TabIndex        =   16
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番設定日"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "F1200121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Const JIGYOBA_FIX$ = "00036003"     '事業場コード(固定値)

'Private Const DELIMIT_CHR$ = CStr(vbTab)    '出力ﾃﾞｰﾀ区切り文字（Tab）
Private Const DELIMIT_CHR$ = ","            '出力ﾃﾞｰﾀ区切り文字（ｶﾝﾏ）

Dim File_Limit      As Long                 '出力ﾌｧｲﾙ毎の書込み限度件数



Private Const ptxYY% = 0                    '設定日　年
Private Const ptxMM% = 1                    '設定日　月
Private Const ptxDD% = 2                    '設定日　日

Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ


Dim HTANA_DATA      As String               'ホスト棚番設定データフルパス
Dim JGYOBA_CODE     As String               '事業場ｺｰﾄﾞ


Dim JGYOBA_FIX     As String                '事業場コード(固定値)



Dim OUT_SYUSI       As Variant              '(ini) 出力用収支 配列
Dim JYOGAI_SOKO     As Variant              '(ini) 除外倉庫 配列

Dim T_JGYOBU()      As String * 1           '対象事業部テーブル


'Private Const LAST_UPDATE_DAY = "ホスト棚番設定データ自動作成[F120012] 2016.02.16 16:15"
Private Const LAST_UPDATE_DAY = "ホスト棚番設定データ自動作成[F120012] 2016.07.13 14:30"


Private Function OUTPUT_Proc(j As Integer) As Integer

Dim sts                     As Integer
Dim ZAIKO_com               As Integer
Dim ITEM_com                As Integer


Dim Location1               As String
Dim Location2               As String
Dim Location3               As String


Dim Ret                     As Long
Dim FileNo                  As Long
Dim FileName                As String       'ﾌｧｲﾙ名
Dim FileSeq                 As Long         'ﾌｧｲﾙ名SEQ 番号
Dim SWk                     As String

Dim c                       As String * 128
Dim Soko_No                 As String * 2

Dim Count                   As Long
Dim Put_Cnt                 As Long

Dim i                       As Long

Dim Skip_Flg                As Boolean
Dim Fsw                     As Boolean

    OUTPUT_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    FileNo = FreeFile
    FileName = HTANA_DATA




    FileSeq = 1

    Ret = InStr(1, Trim(FileName), ".") - 1
    SWk = Left(Trim(FileName), Ret) & "_" & Last_JGYOBU & "_" & Format(FileSeq, "000") & _
          Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    Fsw = True
    DoEvents

    On Error GoTo Error_Proc
    Open (SWk) For Output As FileNo
    Call UniCode_Conv(K3_ITEM.JGYOBU, Last_JGYOBU)

    Call UniCode_Conv(K3_ITEM.ST_SET_DT, "")

    Count = 0
    Put_Cnt = 0
    ITEM_com = BtOpGetGreaterEqual


    lblJGYOBU(j).Caption = Last_JGYOBU

    Do
        DoEvents

        Skip_Flg = False

        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K3_ITEM, Len(K3_ITEM), 3)

        Select Case sts
            Case BtNoErr

                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If



            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "品目マスタ")
                Exit Function
        End Select

        Location1 = ""
        Location2 = ""
        Location3 = ""




''        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
''            Skip_Flg = True
''        Else


        If StrConv(ITEMREC.ST_SET_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text And _
            StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text And _
            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text Then
            
            Skip_Flg = True
        
        Else

'>>>>>>>>>>>>>>>>>  2016.07.13  国内供給区分のチェック追加
'            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Or StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "-" Then
                Skip_Flg = True
'>>>>>>>>>>>>>>>>>  2016.07.13  国内供給区分のチェック追加
            Else



                Call UniCode_Conv(K0_ZAIKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
                ZAIKO_com = BtOpGetGreaterEqual
                sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        If StrConv(ITEMREC.ST_SOKO, vbUnicode) = StrConv(ZAIKOREC.Soko_No, vbUnicode) And _
                            StrConv(ITEMREC.ST_RETU, vbUnicode) = StrConv(ZAIKOREC.Retu, vbUnicode) And _
                            StrConv(ITEMREC.ST_REN, vbUnicode) = StrConv(ZAIKOREC.Ren, vbUnicode) And _
                            StrConv(ITEMREC.ST_DAN, vbUnicode) = StrConv(ZAIKOREC.Dan, vbUnicode) And _
                            Last_JGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode) And _
                            StrConv(ITEMREC.NAIGAI, vbUnicode) = StrConv(ZAIKOREC.NAIGAI, vbUnicode) And _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode) = StrConv(ZAIKOREC.HIN_GAI, vbUnicode) Then
    
    
                            Location1 = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            Location1 = Jyogai_Chk(Location1)
                            If Location1 = "" Then
                                Skip_Flg = True
                            End If
    
                        Else
    
                            Location1 = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                            Location1 = Jyogai_Chk(Location1)
                            If Location1 = "" Then
    
                                Skip_Flg = True
                            Else
    
    
                                Location1 = Location1 & " *"
    
                            End If
    
                        End If
                    Case BtErrEOF
                        Location1 = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
                        Location1 = Jyogai_Chk(Location1)
                        If Location1 = "" Then
    
                            Skip_Flg = True
                        Else
    
                            Location1 = Location1 & " *"
    
                        End If
                    Case Else
                        Call File_Error(sts, ZAIKO_com, "在庫データ")
                        Exit Function
                End Select
    
    
                If Not Skip_Flg Then
    
                    If Fsw Then
                        
                        Print #FileNo,
                        Print #FileNo,
                        
                        Fsw = False
                    End If
    
                    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                    Call UniCode_Conv(K6_ZAIKO.Retu, "")
                    Call UniCode_Conv(K6_ZAIKO.Ren, "")
                    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    
    
                    ZAIKO_com = BtOpGetGreater
    
                    Do
                        DoEvents
    
                        sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
    
                        Select Case sts
                            Case BtNoErr
                                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                                    Exit Do
                                End If
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, ZAIKO_com, "在庫データ")
                                Exit Function
                        End Select
    
    
                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    
                        Select Case sts
                            Case BtNoErr
                                If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                                Else
                                    If Location1 = "" Then
                                        Location1 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Dan, vbUnicode)
                                        Location1 = Jyogai_Chk(Location1)
                                    Else
                                        If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                            StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(ZAIKOREC.Dan, vbUnicode)) = Location1 Then
                                        Else
    
                                            If Location2 = "" Then
                                                Location2 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                            StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                            StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                            StrConv(ZAIKOREC.Dan, vbUnicode)
                                                Location2 = Jyogai_Chk(Location2)
                                            Else
                                                If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                    StrConv(ZAIKOREC.Dan, vbUnicode)) = Location2 Then
                                                Else
                                                    Location3 = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                                    Location3 = Jyogai_Chk(Location3)
                                                    If Location3 <> "" Then
                                                        Exit Do
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
    
                            Case BtErrKeyNotFound
    
                                MsgBox "[倉庫ﾏｽﾀ異常]  Soko_No = " & StrConv(ZAIKOREC.Soko_No, vbUnicode)
    
                                Exit Do
                            Case Else
                                Call File_Error(sts, ZAIKO_com, "在庫データ")
                                Exit Function
                        End Select
    
                        ZAIKO_com = BtOpGetNext
    
                    Loop
                End If
                If Right(Location1, 1) = "*" And Location2 = "" And Location3 = "" Then
                    
                    
                    If Last_JGYOBU = AIRCON Then
                    '在庫なし
                        Location1 = ""
                    End If
                End If
    
                If Last_JGYOBU <> AIRCON Then
                    If Right(Location1, 1) = "*" Then
                        Location1 = Trim(Left(Location1, Len(Location1) - 1))
                    End If
                End If
    
                '棚番設定ﾃﾞｰﾀ出力（出力件数＝棚×ini定義した収支）
                For i = 0 To UBound(OUT_SYUSI)
                    If Put_Cnt = File_Limit Then
                        Close #FileNo
                        FileSeq = FileSeq + 1
    
                        Ret = InStr(1, Trim(FileName), ".") - 1
                        SWk = Left(Trim(FileName), Ret) & "_" & Last_JGYOBU & "_" & Format(FileSeq, "000") & _
                              Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
                        Open (SWk) For Output As FileNo
    
                        Label(7).Caption = SWk
                        DoEvents
    
                        Put_Cnt = 0
                    
                        Print #FileNo,
                        Print #FileNo,
                    
                    End If
    
                    If Len(Trim(Location1)) <> 0 Then
    
                        If Not GetIni(App.EXEName, Left(Location1, 2), App.EXEName, c) Then
                            Location1 = Trim(c) & _
                                        Right(Location1, Len(Location1) - 2)
                        End If
                    End If
    
                    If Len(Trim(Location2)) <> 0 Then
    
                        If Not GetIni(App.EXEName, Left(Location2, 2), App.EXEName, c) Then
                            Location2 = Trim(c) & _
                                        Right(Location2, Len(Location2) - 2)
                        End If
                    End If
    
                    If Len(Trim(Location3)) <> 0 Then
    
                        If Not GetIni(App.EXEName, Left(Location3, 2), App.EXEName, c) Then
                            Location3 = Trim(c) & _
                                        Right(Location3, Len(Location3) - 2)
                        End If
                    End If
    
    
    
                    Print #FileNo, JGYOBA_FIX & DELIMIT_CHR;            '事業場コード（固定値）
                    Print #FileNo, JGYOBA_CODE & DELIMIT_CHR;           '資産管理事業場コード（ini定義）
                                                                        '品目番号
                    Print #FileNo, Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & DELIMIT_CHR;
                    Print #FileNo, OUT_SYUSI(i) & DELIMIT_CHR;          '在庫収支コード（ini定義）
                    Print #FileNo, JGYOBA_CODE & DELIMIT_CHR;           '資産管理在庫収支ｺｰﾄﾞ（=資産管理事業場ｺｰﾄﾞ）
                    Print #FileNo, "00000000" & DELIMIT_CHR;            '補助在庫収支コード
                    Print #FileNo, "0:在庫引当する" & DELIMIT_CHR;      '在庫引当在庫収支区分
                    Print #FileNo, Trim(Location1) & DELIMIT_CHR;       'ロケーション番号１
                    Print #FileNo, Trim(Location2) & DELIMIT_CHR;       'ロケーション番号２
                    Print #FileNo, Trim(Location3) & DELIMIT_CHR;       'ロケーション番号３
                    Print #FileNo, Trim(StrConv(ITEMREC.K_KEITAI, vbUnicode)) & DELIMIT_CHR;    '個装形態コード
                    Print #FileNo, DELIMIT_CHR;                         '出庫担当者コード
'                    Print #FileNo, DELIMIT_CHR;                         '登録ユーザーＩＤ
'                    Print #FileNo, DELIMIT_CHR;                         '登録日付
'                    Print #FileNo, DELIMIT_CHR;                         '登録時刻
'                    Print #FileNo, DELIMIT_CHR;                         '更新ユーザ
'                    Print #FileNo, Format(Now, "YYYY/m/d") & DELIMIT_CHR;   '更新日付
'                    Print #FileNo, Format(Now, "HHMMDD")                '更新時刻
    
    
                    Print #FileNo, "SDCPOS" & DELIMIT_CHR;                  '登録ユーザーＩＤ
                    Print #FileNo, Format(Now, "YYYY/m/d") & DELIMIT_CHR;   '登録日付
                    Print #FileNo, Format(Now, "HHMMDD") & DELIMIT_CHR;     '登録時刻
                    Print #FileNo, "SDCPOS" & DELIMIT_CHR;                  '更新ユーザ
                    Print #FileNo, Format(Now, "YYYY/m/d") & DELIMIT_CHR;   '更新日付
                    Print #FileNo, Format(Now, "HHMMDD") & DELIMIT_CHR;     '更新時刻
    
                    Print #FileNo, "1"                                      '更新時刻
    
    
    
                    Put_Cnt = Put_Cnt + 1
                    lblOUT_CNT(j).Caption = Put_Cnt
                    DoEvents
    
                Next i
            End If
        End If

        Count = Count + 1
        lblin_Cnt(j).Caption = "/" & Format(Count, "#0")
        DoEvents

        ITEM_com = BtOpGetNext

    Loop


    Close #FileNo

    Call Input_UnLock         '画面項目ロック解除
'    Beep
'    MsgBox "「" & FileName & "」は正常に出力されました。"

    OUTPUT_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function

Private Function Jyogai_Chk(pTANABAN As String) As String
'----------------------------------------------------------------------------
'           除外倉庫チェック                        2005/05/16
'
'   引　数　：棚番(倉庫+列+連+段)
'   戻り値　：対象倉庫→引数のまま
'   　　　　：除外倉庫→空文字列
'----------------------------------------------------------------------------
Dim i   As Integer

    Jyogai_Chk = pTANABAN

    For i = 0 To UBound(JYOGAI_SOKO)
        If Left(pTANABAN, 2) = JYOGAI_SOKO(i) Then
            Jyogai_Chk = ""
            Exit For
        End If
    Next i

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1200121.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200121)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200121)


    F1200121.MousePointer = vbDefault

End Sub
                                            'エラーチェック
Private Function Err_Chk() As Integer

Dim i   As Integer


    Err_Chk = True


    For i = ptxYY To ptxDD

        If Text(i).Text = "" Then

            Select Case i
                Case ptxYY
                    Text(i).Text = "0000"
                Case Else
                    Text(i).Text = "00"
            End Select

        Else
            If Not IsNumeric(Text(i).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(i).SetFocus
                Exit Function
            Else
                If i <> ptxYY Then
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End If
    Next i

    Err_Chk = False

End Function




Private Sub Form_Activate()

Dim i   As Integer
    
    
    For i = 0 To UBound(T_JGYOBU)
        Last_JGYOBU = T_JGYOBU(i)
    
        If OUTPUT_Proc(i) Then
            Unload Me
        End If
    
    Next i

    Unload Me

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

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String
Dim sts     As Integer

Dim wkVAL   As Variant

Dim wSTART_DATE  As String * 10
Dim START_DATE  As String * 8

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = Trim(c)
                                '在庫ファイル名取り込み
    If GetIni("FILE", "BU_HTANA_DATA", "SYS", c) Then
        Beep
        MsgBox "ホスト棚番設定データファイル名の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    HTANA_DATA = Trim(c)


'出力用収支・除外倉庫　取込み ######################################################## 2005/05/16 Add ↓
                                '出力用収支 取り込み    2005/05/16
    If GetIni(App.EXEName, "SYUSI", App.EXEName, c) Then
        Beep
        MsgBox "出力用収支の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    OUT_SYUSI = Split(Trim(c), ",", -1)

                                '除外倉庫 取り込み      2005/05/16
    If GetIni(App.EXEName, "JYOGAI", App.EXEName, c) Then
        Beep
        MsgBox "除外倉庫の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    JYOGAI_SOKO = Split(Trim(c), ",", -1)
'################################################################################### 2005/05/16 Add ↑

                                
                                
                                
                                '事業場(固定)名取り込み
    If GetIni(App.EXEName, "JGYOBA_FIX", App.EXEName, c) Then
        Beep
        MsgBox "「事業場(固定)」の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    JGYOBA_FIX = Trim(c)
                                
                                
                                
                                
                                '事業場名取り込み
    If GetIni(App.EXEName, "JGYOBA_CODE", App.EXEName, c) Then
        Beep
        MsgBox "「事業場」の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    JGYOBA_CODE = Trim(c)

                                'ﾌｧｲﾙ書込み限度件数取り込み
    File_Limit = 5000   'ﾃﾞﾌｫﾙﾄ=5000

    If Not GetIni(App.EXEName, "FIL_LIMIT", App.EXEName, c) Then
        If Val(Trim(c)) > 0 Then
            File_Limit = Val(Trim(c))
        End If
    Else                '取得できない場合、デフォルト
'        Beep
'        MsgBox "「ﾌｧｲﾙ書込み限度件数」の取得に失敗しました。処理を中止して下さい。"
'        End
    End If


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 対象事業部
    If GetIni(App.EXEName, "DEF_JGYOBU", App.EXEName, c) Then
        Beep
        MsgBox "「出力対象事業部なし」の取得に失敗しました。処理を中止して下さい。"
        End
    End If

    wkVAL = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkVAL)
        ReDim Preserve T_JGYOBU(0 To i)
        T_JGYOBU(i) = wkVAL(i)
    
    Next i
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 対象日設定
    If GetIni(App.EXEName, "SETDAY", App.EXEName, c) Then
        c = "0"
    End If
    

    F1200121.Caption = LAST_UPDATE_DAY


    wSTART_DATE = DateAdd("d", Val(c), Format(Now, "YYYY/MM/DD"))

    START_DATE = Format(wSTART_DATE, "YYYYMMDD")

    Text(ptxYY).Text = Mid(START_DATE, 1, 4)
    Text(ptxMM).Text = Mid(START_DATE, 5, 2)
    Text(ptxDD).Text = Mid(START_DATE, 7, 2)
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫ﾏｽﾀＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If



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
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '倉庫ﾏｽﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, ZAIKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫ﾏｽﾀ")
        End If
    End If

    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1200121 = Nothing

    End
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


    If KeyCode <> vbKeyReturn Then Exit Sub

    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


