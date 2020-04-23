VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1030101 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷検品ラベル発行"
   ClientHeight    =   7080
   ClientLeft      =   2025
   ClientTop       =   2655
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Caption         =   "時刻指定再印刷"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3465
      TabIndex        =   19
      Top             =   1320
      Width           =   4845
      Begin VB.ListBox List1 
         Height          =   1020
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "印　刷"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   21.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   2310
         TabIndex        =   20
         Top             =   2040
         Width           =   2430
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "伝票番号指定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3465
      TabIndex        =   16
      Top             =   4440
      Width           =   4845
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2730
         MaxLength       =   7
         TabIndex        =   2
         Top             =   480
         Width           =   960
      End
      Begin VB.TextBox Text1 
         Height          =   375
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   1365
         MaxLength       =   7
         TabIndex        =   1
         Top             =   480
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "印　刷"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   21.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2310
         TabIndex        =   3
         Top             =   1080
         Width           =   2430
      End
      Begin VB.Label Label1 
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2415
         TabIndex        =   18
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "伝票番号"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新規分"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
End
Attribute VB_Name = "F1030101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxS_DEN_NO% = 0
Private Const ptxE_DEN_NO% = 1

Private Const Text_Max% = 1

Private Const plstPrint_Now% = 0


'Private Const LAST_UPDATE_DAY$ = "2013.01.23 08:30"
Private Const LAST_UPDATE_DAY$ = "[F103010] 2016.04.26 09:30"


Dim Pri_Name    As Printer

Private wY_SYU_H_POS    As POSBLK
Private wY_SYU_HREC     As Y_SYU_HREC_Tag
Private wK1_Y_SYU_H     As KEY1_Y_SYU_H
Private wK4_Y_SYU_H     As KEY4_Y_SYU_H

Private Function Print_Proc(Mode As Integer, DATA_CNT As Integer) As Integer
'----------------------------------------------------------------------------
'                   印刷処理
'   mode    0:新規処理
'           1:再印刷
'----------------------------------------------------------------------------

Dim lPrinterHandl   As Long         'ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙを取得

Dim sts             As Integer
Dim com             As Integer
Dim wkcom           As Integer
Dim ans             As Integer

Dim sEditWK         As String       '編集ﾜｰｸ
Dim sJis            As String       '漢字変換のﾘﾀｰﾝ
Dim vjis            As String
    
Dim SEQ_NO          As Long
    
Dim DEN_SU          As String
    
Dim SKIP_Flg        As Boolean
    
Dim SV_ID_NO        As String * 7
    
Dim NON_PRINT_Flg   As Boolean
    
Dim PRINT_NOW       As String
    
    
    Print_Proc = True
    
    Call Input_Lock
    
        
    PRINT_NOW = Format(Now, "YYYYMMDDHHMMSS")
    
    
'   印刷開始処理
    PrinterDriver_Start "検品ラベル発行", lPrinterHandl

    SEQ_NO = 0
    SV_ID_NO = ""

    Select Case Mode
        Case 0
        '-------------------------------------  新規印刷指示
            Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
        Case 1
        '-------------------------------------  再印刷指示
            Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Text1(ptxS_DEN_NO).Text)
            Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")
    End Select

    com = BtOpGetGreaterEqual

    Do
    
        DoEvents
        
        SKIP_Flg = False
        
        Select Case Mode
            Case 0
            '-------------------------------------  新規印刷指示
            
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                Select Case sts
                    Case BtNoErr
                    
                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Exit Function
                End Select
            
            
            Case 1
            '-------------------------------------  再印刷指示
        
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
                        If Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                            Exit Do
                        End If
                                    
'                       2016.04.26 無条件印刷とする
'                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
'                            SKIP_Flg = True
'                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Exit Function
                End Select
        
        
        End Select
        
        NON_PRINT_Flg = False
        If Trim(SV_ID_NO) = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
            NON_PRINT_Flg = True
        End If
        SV_ID_NO = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7)
        
        
        If Not SKIP_Flg Then
    
            If Not NON_PRINT_Flg Then
    
    
        '       STX指定
                sEditWK = Chr(&H2)
        '       ﾃﾞｰﾀ送信開始指定
                sEditWK = sEditWK & Chr(&H1B) & "A"
    '2006.12.19
                sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+220"
            
                '伝票番号
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode)
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                '運送会社
                vjis = Kanji_Conv("H", Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)))
                sEditWK = sEditWK & Chr(&H1B) & "H0160" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                '連番
                If Mode = 0 Then
                    SEQ_NO = SEQ_NO + 1
                    
                    sEditWK = sEditWK & Chr(&H1B) & "H0330" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                    sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(SEQ_NO, "#0")
                
                
                End If
                '伝票番号ﾊﾞｰｺｰﾄﾞ
                sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode) & "*"
                sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) & "*"
                
                '得意先ｺｰﾄﾞ
    '''            sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0100" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode))
                '得意先名(送り先)
                vjis = Kanji_Conv("H", StrConv(Trim(Left(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode), 15)), vbWide))
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
    '''            '伝票番号ﾊﾞｰｺｰﾄﾞ
    '''            sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) & "*"
            
                '最大伝票行数の獲得
                Call UniCode_Conv(wK4_Y_SYU_H.ID_NO, Left(Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)), 7) & "99")
                sts = BTRV(BtOpGetLessEqual, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                    
                        If Left(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 7) <> Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
                            DEN_SU = "01"
                        Else
                            DEN_SU = Right(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 2)
                        End If
                    Case BtErrEOF
                        DEN_SU = "01"
                    Case Else
                        Call File_Error(sts, BtOpGetLessEqual, "出荷予定")
                        Exit Function
                End Select
                If Not IsNumeric(DEN_SU) Then
                    DEN_SU = "01"
                End If
                sEditWK = sEditWK & Chr(&H1B) & "H0290" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(DEN_SU, "#0")
                vjis = Kanji_Conv("H", "点")
                sEditWK = sEditWK & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                
                
            
            
    '''            If PRINT_CNT = DATA_CNT Then
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT1"
    '''            Else
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT0"
    '''            End If
                    
                    
                '次レコードの確認
                Call UniCode_Conv(wK1_Y_SYU_H.PRINT_NOW, StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.INS_NOW, StrConv(Y_SYU_HREC.INS_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.DATA_CNT, StrConv(Y_SYU_HREC.DATA_CNT, vbUnicode))
                                
                wkcom = BtOpGetGreater
                                
                Do
                
                    DoEvents
                
                    sts = BTRV(wkcom, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK1_Y_SYU_H, Len(wK1_Y_SYU_H), 1)
                    Select Case sts
                        Case BtNoErr
                            If Mode = 1 Then
                                If Trim(StrConv(wY_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                                    Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                    Exit Do
                                End If
                            End If
                                                
                            If Mode = 0 Then
                                If Trim(StrConv(wY_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                                    Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                    Exit Do
                                End If
                            End If
                            
                            
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> Left(StrConv(wY_SYU_HREC.ID_NO, vbUnicode), 7) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                            Exit Function
                    End Select
                    
                    wkcom = BtOpGetNext
                    
                    
                Loop
                    
                    
                    
                If Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)) <> "" Then
                    sEditWK = sEditWK & Chr(&H1B) & "CT0"
                Else
                    sEditWK = sEditWK & Chr(&H1B) & "CT1"
                
                End If
            
            
        '       指定枚数
                sEditWK = sEditWK & Chr(&H1B) & "Q1"
        
            
        '       ﾃﾞｰﾀ送信終了指定
                sEditWK = sEditWK & Chr(&H1B) & "Z"
        
        '       ETX指定
                sEditWK = sEditWK & Chr(&H3)
            
        '       ﾃﾞｰﾀ送信
                PrinterDriver_Write lPrinterHandl, sEditWK
            End If
        
            '印刷済更新
            If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
                
                Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, PRINT_NOW)
                
                Do
                    Select Case Mode
                        Case 0
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                        Case 1
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                    End Select
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ﾏｽﾀ")
                            Exit Function
                    End Select
                Loop
            
                Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
                Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
                Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
            
                com = BtOpGetGreater
            Else
                com = BtOpGetNext
            
            End If
        Else
            com = BtOpGetNext
        End If
        
    Loop




    '印刷終了処理
    
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_Proc = False


End Function
Private Function Print_RE_Proc() As Integer
'----------------------------------------------------------------------------
'                   最終印刷指示分再印刷処理
'----------------------------------------------------------------------------

Dim lPrinterHandl   As Long         'ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙを取得

Dim sts             As Integer
Dim com             As Integer
Dim wkcom             As Integer
Dim ans             As Integer

Dim sEditWK         As String       '編集ﾜｰｸ
Dim sJis            As String       '漢字変換のﾘﾀｰﾝ
Dim vjis            As String
    
Dim SEQ_NO          As Long
    
Dim DEN_SU          As String
    
Dim SKIP_Flg        As Boolean
    
Dim SV_ID_NO        As String * 7
    
Dim NON_PRINT_Flg   As Boolean
    
Dim SV_PRINT_NOW    As String
    
    Print_RE_Proc = True
    
    Call Input_Lock
    
        
    
    
'   印刷開始処理
    PrinterDriver_Start "検品ラベル発行", lPrinterHandl

    SEQ_NO = 0
    SV_ID_NO = ""

    '最終印刷日時の獲得
'''    sts = BTRV(BtOpGetLast, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
'''    Select Case sts
'''        Case BtNoErr
'''
'''            If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) = "" Then
'''                Print_RE_Proc = False
'''                Exit Function
'''            End If
'''
'''        Case BtErrEOF
'''            Print_RE_Proc = False
'''            Exit Function
'''        Case Else
'''            Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
'''            Exit Function
'''    End Select
'''
'''    SV_PRINT_NOW = StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)
    
    SV_PRINT_NOW = Format(List1(plstPrint_Now).List(List1(plstPrint_Now).ListIndex), "YYYYMMDDHHMMSS")
    
    
    Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, SV_PRINT_NOW)
    Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
    Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")

    com = BtOpGetGreaterEqual

    Do
    
        DoEvents
        
        SKIP_Flg = False
        
            
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
        Select Case sts
            Case BtNoErr
            
                If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> SV_PRINT_NOW Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                Exit Function
        End Select
        
        NON_PRINT_Flg = False
        If Trim(SV_ID_NO) = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
            NON_PRINT_Flg = True
        End If
        SV_ID_NO = Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7)
        
        
        If Not SKIP_Flg Then
    
            If Not NON_PRINT_Flg Then
    
    
        '       STX指定
                sEditWK = Chr(&H2)
        '       ﾃﾞｰﾀ送信開始指定
                sEditWK = sEditWK & Chr(&H1B) & "A"
    '2006.12.19
                sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+220"
            
                '伝票番号
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode)
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode))
                '運送会社
                vjis = Kanji_Conv("H", Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)))
                sEditWK = sEditWK & Chr(&H1B) & "H0160" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                '連番
                SEQ_NO = SEQ_NO + 1
                
                sEditWK = sEditWK & Chr(&H1B) & "H0330" & Chr(&H1B) & "V0030" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(SEQ_NO, "#0")
                
                
                '伝票番号ﾊﾞｰｺｰﾄﾞ
                sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0070" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) & StrConv(Y_SYU_HREC.SEQ_NO, vbUnicode) & "*"
                sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) & "*"
                
                '得意先ｺｰﾄﾞ
    '''            sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0100" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
    '''            sEditWK = sEditWK & Chr(&H1B) & "X21," & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode))
                '得意先名(送り先)
                vjis = Kanji_Conv("H", StrConv(Trim(Left(StrConv(Y_SYU_HREC.OKURISAKI, vbUnicode), 15)), vbWide))
                sEditWK = sEditWK & Chr(&H1B) & "H0010" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
    '''            '伝票番号ﾊﾞｰｺｰﾄﾞ
    '''            sEditWK = sEditWK & Chr(&H1B) & "H060" & Chr(&H1B) & "V0130" & Chr(&H1B) & "L0101"
    '''            sEditWK = sEditWK & Chr(&H1B) & "D102040" & "*" & Trim(StrConv(Y_SYU_HREC.MUKE_CODE, vbUnicode)) & "*"
            
                '最大伝票行数の獲得
                Call UniCode_Conv(wK4_Y_SYU_H.ID_NO, Left(Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)), 7) & "99")
                sts = BTRV(BtOpGetLessEqual, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 4)
                Select Case sts
                    Case BtNoErr
                    
                        If Left(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 7) <> Left(Trim(StrConv(Y_SYU_HREC.ID_NO, vbUnicode)), 7) Then
                            DEN_SU = "01"
                        Else
                            DEN_SU = Right(Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)), 2)
                        End If
                    Case BtErrEOF
                        DEN_SU = "01"
                    Case Else
                        Call File_Error(sts, BtOpGetLessEqual, "出荷予定")
                        Exit Function
                End Select
                If Not IsNumeric(DEN_SU) Then
                    DEN_SU = "01"
                End If
                sEditWK = sEditWK & Chr(&H1B) & "H0290" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0102" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "X21," & Format(DEN_SU, "#0")
                vjis = Kanji_Conv("H", "点")
                sEditWK = sEditWK & Chr(&H1B) & "L0101" & Chr(&H1B) & "P00"
                sEditWK = sEditWK & Chr(&H1B) & "K2H" & vjis
                
                
            
            
    '''            If PRINT_CNT = DATA_CNT Then
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT1"
    '''            Else
    '''                sEditWK = sEditWK & Chr(&H1B) & "CT0"
    '''            End If
                    
                    
                '次レコードの確認
                Call UniCode_Conv(wK1_Y_SYU_H.PRINT_NOW, StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.INS_NOW, StrConv(Y_SYU_HREC.INS_NOW, vbUnicode))
                Call UniCode_Conv(wK1_Y_SYU_H.DATA_CNT, StrConv(Y_SYU_HREC.DATA_CNT, vbUnicode))
                                
                wkcom = BtOpGetGreater
                                
                Do
                
                    DoEvents
                
                    sts = BTRV(wkcom, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK1_Y_SYU_H, Len(wK1_Y_SYU_H), 1)
                    Select Case sts
                        Case BtNoErr
                                            
                            If Trim(StrConv(wY_SYU_HREC.PRINT_NOW, vbUnicode)) <> SV_PRINT_NOW Then
                                Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                                Exit Do
                            End If
                            
                            
                            If Left(StrConv(Y_SYU_HREC.ID_NO, vbUnicode), 7) <> Left(StrConv(wY_SYU_HREC.ID_NO, vbUnicode), 7) Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Call UniCode_Conv(wY_SYU_HREC.ID_NO, "")
                            Exit Do
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                            Exit Function
                    End Select
                    
                    wkcom = BtOpGetNext
                    
                    
                Loop
                    
                    
                    
                If Trim(StrConv(wY_SYU_HREC.ID_NO, vbUnicode)) <> "" Then
                    sEditWK = sEditWK & Chr(&H1B) & "CT0"
                Else
                    sEditWK = sEditWK & Chr(&H1B) & "CT1"
                
                End If
            
            
        '       指定枚数
                sEditWK = sEditWK & Chr(&H1B) & "Q1"
        
            
        '       ﾃﾞｰﾀ送信終了指定
                sEditWK = sEditWK & Chr(&H1B) & "Z"
        
        '       ETX指定
                sEditWK = sEditWK & Chr(&H3)
            
        '       ﾃﾞｰﾀ送信
                PrinterDriver_Write lPrinterHandl, sEditWK
            End If
        End If
            
        com = BtOpGetNext
        
            
    Loop




    '印刷終了処理
    
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_RE_Proc = False


End Function

Private Sub Command_Click(Index As Integer)

Dim sts         As Integer
Dim i           As Integer
Dim Tana_Cnt    As Long
Dim Yn          As Integer
    
    
    
    Select Case Index
        
        
        
        
        Case 11                             '「終了」
            Unload Me
        Case Else
            Beep
    End Select
    
    Exit Sub
    
    
    
End Sub


Private Sub Command1_Click(Index As Integer)
'----------------------------------------------------------------------------
'                   新規印刷の指示
'
'----------------------------------------------------------------------------



Dim DATA_CNT    As Integer
Dim Yn          As Integer


 '''           DATA_CNT = Print_Cnt_Proc(0)
 '''           If DATA_CNT < 0 Then
 '''               Unload Me
 '''           End If
        
 '''           Yn = MsgBox("検品ラベルは「" & StrConv(Format(DATA_CNT, "#,##0"), vbWide) & "」枚発行されます。宜しいですか？", vbYesNo, "確認入力")
        
            
    Select Case Index
        Case 0
            
            Yn = MsgBox("「検品ラベル」新規印刷を行いますか？", vbYesNo, "確認入力")
        
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_Proc(0, DATA_CNT) Then
                    Unload Me
                End If
        
        
        
            End If

        Case 1

            If List1(plstPrint_Now).ListIndex < 0 Then
                MsgBox "指定行を選択してください"
                
                List1(plstPrint_Now).SetFocus
                List1(plstPrint_Now).ListIndex = 0
                
                Exit Sub
            End If
            
            
            
            'Yn = MsgBox("「検品ラベル」最終印刷指示分再印刷を行いますか？", vbYesNo, "確認入力")
            Yn = MsgBox("「検品ラベル」時刻指定分　再印刷を行いますか？", vbYesNo, "確認入力")          '2012.12.27 修正    M.T
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_RE_Proc() Then
                    Unload Me
                End If
        
            Else
            
        
            End If

    End Select

ErrHandler:

End Sub

Private Sub Command2_Click()
Dim DATA_CNT    As Integer
Dim Yn          As Integer


'''            DATA_CNT = Print_Cnt_Proc(1)
'''            If DATA_CNT < 0 Then
'''                Unload Me
'''            End If
        
'''            Yn = MsgBox("検品ラベルは「" & StrConv(Format(DATA_CNT, "#,##0"), vbWide) & "」枚発行されます。宜しいですか？", vbYesNo, "確認入力")
            
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.12.27  追加    M.T
            If Trim(Text1(ptxS_DEN_NO)) = "" And Trim(Text1(ptxE_DEN_NO)) = "" Then
                MsgBox "開始伝票番号の指定が不正　→　印刷不能!", vbExclamation
                Text1(ptxS_DEN_NO).SetFocus
                Call Text1_GotFocus(ptxS_DEN_NO)
                Exit Sub
            End If
            
            If Trim(Text1(ptxE_DEN_NO)) = "" Then
                Text1(ptxE_DEN_NO) = Text1(ptxS_DEN_NO)
            End If
            
            If Trim(Text1(ptxS_DEN_NO)) > Trim(Text1(ptxE_DEN_NO)) Then
                MsgBox "伝票番号の指定が不正　→　印刷不能！", vbExclamation
                Text1(ptxE_DEN_NO).SetFocus
                Call Text1_GotFocus(ptxE_DEN_NO)
                Exit Sub
            End If
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>     ここまで
            
            Yn = MsgBox("「検品ラベル」再印刷を行いますか？", vbYesNo, "確認入力")
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_Proc(1, DATA_CNT) Then
                    Unload Me
                End If
                
                
                Text1(ptxE_DEN_NO) = ""
            Else
            
                Text1(ptxS_DEN_NO).SetFocus                     '2012.12.27  追加    M.T
                Call Text1_GotFocus(ptxS_DEN_NO)                '2012.12.27  追加    M.T
            
            End If

ErrHandler:

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
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    F1030101.Caption = F1030101.Caption & LAST_UPDATE_DAY           '2016.04.26
    

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    
    
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    
                                
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If Y_SYU_H_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ
    If wY_SYU_H_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                
                                
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Exit For
        End If
    Next
                                    
    '印刷済みを取り込む
    If Print_Re_Set_Proc() Then
        Unload Me
    End If
                                
                                
    Command1(0).SetFocus
End Sub

Private Sub Form_Unload(CANCEL As Integer)

Dim sts         As Integer
Dim Wk_Printer  As Printer
                                            
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Pri_Name.DeviceName) Then
            SetWindowsDefaultPrinter Wk_Printer.DeviceName, Wk_Printer.DriverName, Wk_Printer.Port
            Exit For
        End If
    Next
                                            
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
    
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), wK4_Y_SYU_H, Len(wK4_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1030101 = Nothing


    End
End Sub
Private Function Print_Cnt_Proc(Mode As Integer) As Long
'----------------------------------------------------------------------------
'                   印刷枚数のカウント
'   mode    0:新規分印刷
'           1:再印刷
'
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim DATA_CNT    As Long


    Print_Cnt_Proc = True

    DATA_CNT = 0



    Select Case Mode
        Case 0
        '---------------------------------------------新規分
            Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")


            com = BtOpGetGreaterEqual
        
        
            Do
            
                DoEvents
                
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
                Select Case sts
                    Case BtNoErr
                                            
                        If Trim(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)) <> "" Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Exit Function
                End Select
            
                DATA_CNT = DATA_CNT + 1
            
                com = BtOpGetNext
            
            Loop



        Case 1
        '---------------------------------------------再印刷
    
            Call UniCode_Conv(K0_Y_SYU_H.DEN_NO, Trim(Text1(ptxS_DEN_NO).Text))
            Call UniCode_Conv(K0_Y_SYU_H.SEQ_NO, "")


            com = BtOpGetGreaterEqual
        
        
            Do
            
                DoEvents
                
                sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                Select Case sts
                    Case BtNoErr
                                            
                        If Trim(StrConv(Y_SYU_HREC.DEN_NO, vbUnicode)) > Trim(Text1(ptxE_DEN_NO).Text) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                        Exit Function
                End Select
            
                DATA_CNT = DATA_CNT + 1
            
                com = BtOpGetNext
            
            Loop
    
    
    
    End Select











    Print_Cnt_Proc = DATA_CNT

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1030101.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030101)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030101)


    F1030101.MousePointer = vbDefault

End Sub


Private Function isWindowsNT() As Boolean
  isWindowsNT = IIf(GetVersion() And &H80000000, False, True)
End Function
Private Sub SetWindowsDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal Port As String)
  Dim param As String
  param = DeviceName & "," & DriverName & "," & Port
  WriteProfileString "windows", "device", param
  If isWindowsNT Then
    'Windows NT/2000
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal 0&
  Else
    'Windows 95/98/Me
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal "windows"
  End If
'  Printer.EndDoc
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim i       As Integer

Dim W_Str   As String                                               ' 2012.12.27  追加    M.T

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<    2012.12.27  追加    M.T
    Select Case Index
        Case ptxS_DEN_NO        '開始伝票番号
            If Not IsNumeric(Text1(Index)) Then
                MsgBox "数値エラー！", vbExclamation
                Text1(Index).SetFocus
                Call Text1_GotFocus(Index)
                Exit Sub
            End If
            
            Call Numeric_Check(EDIT_ONLY, Text1(Index).MaxLength, 0, NEGA_DIS, ZSUP_DIS, COMA_DIS, Text1(Index), W_Str)
            Text1(Index) = W_Str
'2013.01.24            If Trim(Text1(ptxE_DEN_NO)) = "" Then
                Text1(ptxE_DEN_NO) = W_Str
'2013.01.24            End If
        Case ptxE_DEN_NO        '終了伝票番号
            Call Numeric_Check(EDIT_ONLY, Text1(Index).MaxLength, 0, NEGA_DIS, ZSUP_DIS, COMA_DIS, Text1(Index), W_Str)
            Text1(Index) = W_Str
        
        Case Else
        
        
    End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>     ここまで
    Call Tab_Ctrl(Shift)     '移動


End Sub
Function Kanji_Conv(psPara As String, psSiftJis As String) As String
Rem ｼﾌﾄJISｺｰﾄﾞからJISｺｰﾄﾞへ変換
'   psPara      :   H=HEX   B=Binary
'   psSiftJis   :   ｼﾌﾄJISｺｰﾄﾞ

Dim i As Integer    '桁数のﾘﾀｰﾝｺｰﾄﾞ
Dim vConv           'ﾜｰｸ変数
Dim vHex            '4ﾊﾞｲﾄのｼﾌﾄJISｺｰﾄﾞに変換のﾘﾀｰﾝｺｰﾄﾞ
Dim vUpByte         '上位2ﾊﾞｲﾄを1ﾊﾞｲﾄに変換のﾘﾀｰﾝｺｰﾄﾞ
Dim vDownByte       '下位2ﾊﾞｲﾄを1ﾊﾞｲﾄに変換のﾘﾀｰﾝｺｰﾄﾞ
    
    vConv = ""                                    'ﾜｰｸ変数の初期化
    For i = 1 To Len(psSiftJis)                   '桁数分繰り返す
        vHex = Hex(Asc(Mid$(psSiftJis, i, 1)))    '４ﾊﾞｲﾄのｼﾌﾄJISｺｰﾄﾞに変換
        If vHex = "20" Then
           Exit For
        End If
        vUpByte = Val("&h" + Mid$(vHex, 1, 2))    '上位２ﾊﾞｲﾄを１ﾊﾞｲﾄに変換
        vDownByte = Val("&h" + Mid$(vHex, 3, 2))  '下位２ﾊﾞｲﾄを１ﾊﾞｲﾄに変換
        If vUpByte >= &HE0 Then                   '上位１ﾊﾞｲﾄがＥ０hの場合の処理
           vUpByte = vUpByte - &H40
        End If
        vUpByte = (vUpByte - &H81) * 2 + &H21
        If vDownByte > &H7F Then                  '下位１ﾊﾞｲﾄが８０h以上の処理
           vDownByte = vDownByte - 1
        End If
        If vDownByte > &H9D Then                  '下位１ﾊﾞｲﾄが９Ｅh以上の処理
           vUpByte = vUpByte + 1
           vDownByte = vDownByte - (&H9E - &H21)
        Else
           vDownByte = vDownByte - (&H40 - &H21)  '下位１ﾊﾞｲﾄが９Ｄ以下の処理
        End If
        Select Case psPara
               Case "H"
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ﾜｰｸ変数に足し込む
               Case "B"
                    vConv = vConv + Chr$(vUpByte) + Chr$(vDownByte)  'ﾜｰｸ変数に足し込む
               Case Else
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ﾜｰｸ変数に足し込む
        End Select
    Next i
    Kanji_Conv = vConv

End Function

Function wY_SYU_H_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ  ＯＰＥＮ
'*
'*      引  数:Open Mode(Btrieve参照)
'*      戻り値:false 正常
'*             true  異常
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String
    
    wY_SYU_H_Open = True
                                            '出荷予定データフルパス取込み
    sts = GetIni("FILE", Y_SYU_H_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [Y_SYU_H]読み込みエラー ")
        Exit Function
    End If
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wY_SYU_H_POS, wY_SYU_HREC, Len(wY_SYU_HREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case Else
                Call File_Error(sts, BtOpOpen, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
                Exit Function
        End Select
    Loop
    wY_SYU_H_Open = False
End Function


Private Function Print_Re_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   再印刷選択用の日付セット
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim svPRINT_NOW     As String * 14


    Print_Re_Set_Proc = True

    
    Call UniCode_Conv(K1_Y_SYU_H.PRINT_NOW, "19900101000000")
    Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "")
    Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "")
    
    com = BtOpGetGreater

    List1(plstPrint_Now).Clear

    svPRINT_NOW = ""
    
    Do
    
        DoEvents
        
        sts = BTRV(com, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K1_Y_SYU_H, Len(K1_Y_SYU_H), 1)
        Select Case sts
            Case BtNoErr
                                    
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
                Exit Function
        End Select
    
        If Trim(svPRINT_NOW) = "" Then
            svPRINT_NOW = StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode)
        
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "zzzzzzzzzzzzzz")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "zzzzz")
        
            
            List1(plstPrint_Now).AddItem Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 7, 2) & " " & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 9, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 11, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 13, 2)

        End If
    
        If svPRINT_NOW <> StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode) Then
        
            Call UniCode_Conv(K1_Y_SYU_H.INS_NOW, "zzzzzzzzzzzzzz")
            Call UniCode_Conv(K1_Y_SYU_H.DATA_CNT, "zzzzz")
        
            List1(plstPrint_Now).AddItem Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 7, 2) & " " & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 9, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 11, 2) & ":" & _
                                            Mid(StrConv(Y_SYU_HREC.PRINT_NOW, vbUnicode), 13, 2)
        
        End If
    
        com = BtOpGetGreater
    
    Loop

    If List1(plstPrint_Now).ListCount = 0 Then
        List1(plstPrint_Now).AddItem "再印刷対象無し"
        Frame2.Enabled = False
    End If


    Print_Re_Set_Proc = False

End Function
