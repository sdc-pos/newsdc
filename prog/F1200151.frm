VERSION 5.00
Begin VB.Form F1200151 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ホスト棚番設定データ作成"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "変更履歴"
      Height          =   255
      Index           =   14
      Left            =   1470
      TabIndex        =   36
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GLICS_TANA1-3 と差異がある場合のみ出力(2009.11.06)"
      Height          =   255
      Left            =   2835
      TabIndex        =   35
      Top             =   5160
      Width           =   8040
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2835
      TabIndex        =   34
      Top             =   4800
      Width           =   8040
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2835
      TabIndex        =   33
      Top             =   4440
      Width           =   8040
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2835
      TabIndex        =   32
      Top             =   4080
      Width           =   8040
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2835
      TabIndex        =   31
      Top             =   3720
      Width           =   8040
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2835
      TabIndex        =   30
      Top             =   3360
      Width           =   8040
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "出力先"
      Height          =   255
      Index           =   13
      Left            =   1470
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "分割件数"
      Height          =   255
      Index           =   12
      Left            =   1470
      TabIndex        =   28
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "除外"
      Height          =   255
      Index           =   11
      Left            =   1470
      TabIndex        =   27
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "収支"
      Height          =   255
      Index           =   10
      Left            =   1470
      TabIndex        =   26
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "事業場"
      Height          =   255
      Index           =   9
      Left            =   1470
      TabIndex        =   25
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   4560
      TabIndex        =   24
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   4560
      TabIndex        =   23
      Top             =   2940
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
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
      Left            =   3240
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
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1200151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const JIGYOBA_FIX$ = "00036003"     '事業場コード(固定値)

'Private Const DELIMIT_CHR$ = CStr(vbTab)    '出力ﾃﾞｰﾀ区切り文字（Tab）
Private Const DELIMIT_CHR$ = ","            '出力ﾃﾞｰﾀ区切り文字（ｶﾝﾏ）

Dim File_Limit      As Long                 '出力ﾌｧｲﾙ毎の書込み限度件数



Private Const ptxYY% = 0                    '設定日　年
Private Const ptxMM% = 1                    '設定日　月
Private Const ptxDD% = 2                    '設定日　日

Private Const Text_Max% = 2                 '画面項目別最大ｲﾝﾃﾞｯｸｽ


Dim HTANA_DATA      As String               'ホスト棚番設定データフルパス
Dim JGYOBA_CODE     As String               '事業場ｺｰﾄﾞ

Dim OUT_SYUSI       As Variant              '(ini) 出力用収支 配列
Dim JYOGAI_SOKO     As Variant              '(ini) 除外倉庫 配列

'Private Const LAST_UPDATE_DAY$ = "(F120015 2009.11.09 17:30)"
Private Const LAST_UPDATE_DAY$ = "(F120015 2016.07.13 16:00)"

Private Function OUTPUT_Proc() As Integer

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
    Label(7).Caption = SWk
    
    Fsw = True
    DoEvents

    On Error GoTo Error_Proc
    Open (SWk) For Output As FileNo
    Call UniCode_Conv(K3_ITEM.JGYOBU, Last_JGYOBU)

''    Call UniCode_Conv(K3_ITEM.ST_SET_DT, Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text)
    Call UniCode_Conv(K3_ITEM.ST_SET_DT, "")

    Count = 0
    Put_Cnt = 0
    ITEM_com = BtOpGetGreaterEqual

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


If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) = "AMC34K-Q00" Then
    Debug.Print
End If


        Location1 = ""
        Location2 = ""
        Location3 = ""




''        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
''            Skip_Flg = True
''        Else


        If Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text = "99999999" Then     '2008.03.05
            If Not IsNumeric(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) Then
                Skip_Flg = True
            Else
                If CLng(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) <> 99999999 Then
                    Skip_Flg = True
                End If
            End If
        End If


        If Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text = "99999999" Then     '2008.03.05
        Else
            If StrConv(ITEMREC.ST_SET_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text And _
                StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text And _
                StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Text(ptxYY).Text & Text(ptxMM).Text & Text(ptxDD).Text Then
                
                Skip_Flg = True
            
            End If
        End If
        
        
        
'>>>>>>>>   2016.07.13 国内供給区分のチェック
        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "-" Then
            Skip_Flg = True
        End If
'>>>>>>>>   2016.07.13 国内供給区分のチェック
        
        
        If Not Skip_Flg Then

            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Skip_Flg = True
            Else



                Call UniCode_Conv(K0_ZAIKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                '2009.11.09
                Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
                
                
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
    
    
'                    If i = 0 Then   '2008.03.05    2008.04.02 DEL
'                        If Not IsNumeric(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) Then
'                        Else
'                            If CLng(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) = 99999999 Then
'                                Location1 = "*" & Location1
'                            End If
'                        End If
'                    End If
    
                    If i = 0 Then   '2008.04.02
                        
                        
                        If StrConv(ITEMREC.H_TANA_MAKE, vbUnicode) = vbNullChar Then
                            Debug.Print
                        Else
                            Location1 = Trim(StrConv(ITEMREC.H_TANA_MAKE, vbUnicode)) & Location1
                        End If
                    End If
    
    
                    'LOCATIONチェック
    
                    If Trim(Location1) = Trim(StrConv(ITEMREC.GLICS1_TANA, vbUnicode)) And _
                        Trim(Location2) = Trim(StrConv(ITEMREC.GLICS2_TANA, vbUnicode)) And _
                        Trim(Location3) = Trim(StrConv(ITEMREC.GLICS3_TANA, vbUnicode)) Then
                    Else
                    'ﾛｹｰｼｮﾝ1、ﾛｹｰｼｮﾝ2、ﾛｹｰｼｮﾝ3の異なるもののみ出力
    
    
    
    
    
    
                        Print #FileNo, JIGYOBA_FIX & DELIMIT_CHR;           '事業場コード（固定値）
                        Print #FileNo, JGYOBA_CODE & DELIMIT_CHR;           '資産管理事業場コード（ini定義）
                                                                            '品目番号
                        Print #FileNo, Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) & DELIMIT_CHR;
                        Print #FileNo, OUT_SYUSI(i) & DELIMIT_CHR;          '在庫収支コード（ini定義）
                        Print #FileNo, JGYOBA_CODE & DELIMIT_CHR;           '資産管理在庫収支ｺｰﾄﾞ（=資産管理事業場ｺｰﾄﾞ）
                        Print #FileNo, "00000000" & DELIMIT_CHR;            '補助在庫収支コード
                        
                        
                        
                        If i = 0 Then       '2008.12.09
                            Print #FileNo, "0:在庫引当する" & DELIMIT_CHR;  '在庫引当在庫収支区分
                        Else
                            Print #FileNo, "1:在庫引当しない" & DELIMIT_CHR;  '在庫引当在庫収支区分
                        End If
                        
                        Print #FileNo, Trim(Location1) & DELIMIT_CHR;       'ロケーション番号１
                        Print #FileNo, Trim(Location2) & DELIMIT_CHR;       'ロケーション番号２
                        Print #FileNo, Trim(Location3) & DELIMIT_CHR;       'ロケーション番号３
                        Print #FileNo, Trim(StrConv(ITEMREC.K_KEITAI, vbUnicode)) & DELIMIT_CHR;    '個装形態コード
                        Print #FileNo, DELIMIT_CHR;                         '出庫担当者コード
'                        Print #FileNo, "SDCPOS" & DELIMIT_CHR;                  '登録ユーザーＩＤ
                        Print #FileNo, DELIMIT_CHR;                         '登録ユーザーＩＤ
'                        Print #FileNo, Format(Now, "YYYY/m/d") & DELIMIT_CHR;   '登録日付
                        Print #FileNo, DELIMIT_CHR;                             '登録日付
'                        Print #FileNo, Format(Now, "HHMMDD") & DELIMIT_CHR;    '登録時刻
                        Print #FileNo, DELIMIT_CHR;                             '登録時刻
                        Print #FileNo, "SDCPOS" & DELIMIT_CHR;                  '更新ユーザ
                        Print #FileNo, Format(Now, "YYYY/MM/DD") & DELIMIT_CHR;   '更新日付
                        Print #FileNo, Left(Format(Now, "HH:MM:DD"), 5) & DELIMIT_CHR;    '更新時刻
        
'                        Print #FileNo, "1"                                      '更新時刻
                        Print #FileNo,                              '更新時刻
        
        
                        Put_Cnt = Put_Cnt + 1
                        Label(8).Caption = Put_Cnt
                        DoEvents
                    End If
                Next i
            End If
        End If

        Count = Count + 1
        Label(6).Caption = "/" & Format(Count, "#0")
        DoEvents

        ITEM_com = BtOpGetNext

    Loop


    Close #FileNo

    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "「" & FileName & "」は正常に出力されました。"

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

    F1200151.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200151)


    F1200151.MousePointer = vbDefault

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


Private Sub Command_Click(Index As Integer)
Dim ans As Integer

    Select Case Index
        Case 7                              'データ出力
            If Err_Chk() Then
                Exit Sub
            End If

            Beep
            ans = MsgBox("「ホスト棚番設定データ」出力しますか？", vbYesNo + vbQuestion, "確認入力")

            If ans = vbYes Then
                If OUTPUT_Proc Then
                    Unload Me
                End If
            End If

            Text(ptxYY).SetFocus

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

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String
Dim sts     As Integer

Dim strWork As String   '2008.12.09


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
    If GetIni("FILE", "HTANA_DATA", "SYS", c) Then
        Beep
        MsgBox "ホスト棚番設定データファイル名の取得に失敗しました。処理を中止して下さい。"
        End
    End If
    HTANA_DATA = Trim(c)
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の取得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1200151.Caption = "ホスト棚番設定データ作成（" + RTrim(JGYOBU_T(i).NAME) + ")" & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)


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



    '2008.12.09
    lblINI(0).Caption = JGYOBA_CODE
    '2008.12.09
    strWork = ""
    For i = 0 To UBound(OUT_SYUSI)
        strWork = strWork & OUT_SYUSI(i) & ","
    Next i
    strWork = Left(strWork, Len(strWork) - 1)
    lblINI(1).Caption = strWork
    '2008.12.09
    strWork = ""
    For i = 0 To UBound(JYOGAI_SOKO)
        strWork = strWork & JYOGAI_SOKO(i) & ","
    Next i
    strWork = Left(strWork, Len(strWork) - 1)
    lblINI(2).Caption = strWork
    '2008.12.09
    lblINI(3).Caption = File_Limit
    '2008.12.09
    lblINI(4).Caption = HTANA_DATA





    '2008.12.09
    Text(ptxYY).Text = Mid(DateAdd("d", -1, Format(Now, "YYYY/MM/DD")), 1, 4)
    '2008.12.09
    Text(ptxMM).Text = Mid(DateAdd("d", -1, Format(Now, "YYYY/MM/DD")), 6, 2)
    '2008.12.09
    Text(ptxDD).Text = Mid(DateAdd("d", -1, Format(Now, "YYYY/MM/DD")), 9, 2)


    Text(ptxYY).SetFocus


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

    Set F1200151 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1200151.Caption = "ホスト棚番設定データ作成（" + RTrim(JGYOBU_T(Index).NAME) + ")" & LAST_UPDATE_DAY
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


    If KeyCode <> vbKeyReturn Then Exit Sub

    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


