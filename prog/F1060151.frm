VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form F1060151 
   BackColor       =   &H00FFFFFF&
   Caption         =   "作業監視モニター"
   ClientHeight    =   6915
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   12195
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
   ScaleHeight     =   6915
   ScaleWidth      =   12195
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   9
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   8
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   7
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   6
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Width           =   372
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000B&
      Height          =   360
      Index           =   5
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5280
      Width           =   732
   End
   Begin VB.CommandButton Command 
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
      Caption         =   "最 新"
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
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   4095
      Left            =   1440
      OleObjectBlob   =   "F1060151.frx":0000
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label lblTotal 
      Caption         =   "Label1"
      Height          =   495
      Left            =   10035
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblInv 
      Caption         =   "Label1"
      Height          =   495
      Left            =   10035
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblCan 
      Caption         =   "Label1"
      Height          =   495
      Left            =   10035
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "分現在"
      Height          =   255
      Index           =   10
      Left            =   10080
      TabIndex        =   17
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "時"
      Height          =   255
      Index           =   9
      Left            =   9240
      TabIndex        =   16
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   8
      Left            =   8280
      TabIndex        =   15
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   14
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   13
      Top             =   5400
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
      TabIndex        =   12
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1060151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxDATE_YY% = 5           '現在　年
Private Const ptxDATE_MM% = 6           '現在　月
Private Const ptxDATE_DD% = 7           '現在　日
Private Const ptxTIME_HH% = 8           '現在　時
Private Const ptxTIME_MM% = 9           '現在　分

Dim Y_SYUKA     As New XArrayDB

Private Const Min_Row% = 1              '最小行数
'Private Const Max_Row& = 2000           '最大行数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 3             '最大列数


Private Const ColCyu_Kbn% = 0           '列　注文区分
Private Const ColSoko_Name% = 1         '列　倉庫名称
Private Const ColDen_Su% = 2            '列　出庫済み数／出荷予定数
Private Const ColProgress% = 3          '列　進捗％


Private Const LAST_UPDATE_DAY$ = "(F106015 2010.08.30 09:00)"


Private Function List_Dsp_Proc() As Integer
    
Dim com             As Integer
Dim sts             As Integer
Dim i               As Integer

Dim Save_Cyu_Kbn    As String * 1
Dim Save_Soko_No    As String * 2

Dim Syo_Kei_ALL     As Long
Dim Syo_Kei_Sumi    As Long
Dim Cyu_Kei_ALL     As Long
Dim Cyu_Kei_Sumi    As Long
Dim Sou_Kei_ALL     As Long
Dim Sou_Kei_Sumi    As Long

Dim Ritu            As Double


Dim Row             As Integer
    
Dim SKIP_F          As Boolean
Dim FAST_F          As Boolean
    
    
Dim Total_Cnt       As Long
Dim Total_Can       As Long
Dim Total_Inv       As Long
    
    
    List_Dsp_Proc = True
    
    Call Input_Lock
                                    'テーブルリセット
    Set Y_SYUKA = Nothing
    
    Row = 0
    
   
    Syo_Kei_ALL = 0
    Syo_Kei_Sumi = 0
    Cyu_Kei_ALL = 0
    Cyu_Kei_Sumi = 0
    Sou_Kei_ALL = 0
    Sou_Kei_Sumi = 0
    
    
Total_Cnt = 0
Total_Inv = 0
Total_Can = 0
    
    
    Call UniCode_Conv(K6_Y_SYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, "")
    Call UniCode_Conv(K6_Y_SYU.HTANABAN, "")
    Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
    Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    
    com = BtOpGetGreaterEqual
    FAST_F = True
    Do
        
        DoEvents
        
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
        Select Case sts
            Case BtNoErr
                                            '事業部ブレーク
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If

Total_Cnt = Total_Cnt + 1
lblTotal.Caption = Total_Cnt
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定データ")
                List_Dsp_Proc = SYS_ERR
                Exit Function
        End Select
                                        
        'ｷｬﾝｾﾙ/備考の有無のﾁｪｯｸ 2007.01.15
        SKIP_F = False
        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
        sts = BTRV(BtOpGetEqual, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
        Select Case sts
            Case BtNoErr
            
                If Trim(StrConv(Y_SYU_HREC.CANCEL_F, vbUnicode)) = "1" Then
Total_Can = Total_Can + 1
lblCan.Caption = Total_Can
                    SKIP_F = True
                End If
            
                If Left(StrConv(Y_SYU_HREC.INPUT_BIKOU, vbUnicode), 1) < " " Then
                Else
                    If Trim(StrConv(Y_SYU_HREC.INPUT_BIKOU, vbUnicode)) <> "" Then
                        
                        
                        
If Not SKIP_F Then
    Total_Can = Total_Can + 1
    lblCan.Caption = Total_Can
End If
                        
                        SKIP_F = True
                    End If
                End If
            
            
            
            Case BtErrKeyNotFound
                SKIP_F = True           '異常ﾃﾞｰﾀ
Total_Inv = Total_Inv + 1
lblInv.Caption = Total_Inv
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                List_Dsp_Proc = SYS_ERR
                Exit Function
        End Select
                                        
                                        
        If Not SKIP_F Then
            If FAST_F Then
                Save_Cyu_Kbn = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
                Save_Soko_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
                FAST_F = False
            End If
                                            
                                             
            If Save_Cyu_Kbn <> StrConv(Y_SYUREC.CYU_KBN, vbUnicode) Then
                
                

                
                Row = Row + 1
                Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                            
                Select Case Save_Cyu_Kbn
                    Case CYU_KBN_TUK                '月切り
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_1
                    Case CYU_KBN_SPO                '
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_2
                    Case CYU_KBN_HJU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_3
                    Case CYU_KBN_TOK
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_4
                    Case CYU_KBN_BOU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_E
                End Select
                                        
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko_No)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Y_SYUKA(Row, ColSoko_Name) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
'2010.08.26
'                        Y_SYUKA(Row, ColSoko_Name) = "？？"
                        Y_SYUKA(Row, ColSoko_Name) = Save_Soko_No


Debug.Print "1 Save_Soko_No=" & Save_Soko_No

'2010.08.26
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        List_Dsp_Proc = SYS_ERR
                        Exit Function
                End Select
                                                   
                Y_SYUKA(Row, ColDen_Su) = Format(Syo_Kei_Sumi, "#,##0") & "/" & Format(Syo_Kei_ALL, "#,##0")
                If Syo_Kei_ALL <> 0 Then
                    Ritu = Syo_Kei_Sumi / Syo_Kei_ALL
                Else
                    Ritu = 0
                End If
                Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                                
                Syo_Kei_Sumi = 0
                Syo_Kei_ALL = 0
                Save_Soko_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
                                            
                                            
                Row = Row + 1
                Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                            
                Select Case Save_Cyu_Kbn
                    Case CYU_KBN_TUK                '月切り
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_1 & "　計"
                    Case CYU_KBN_SPO                '
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_2 & "　計"
                    Case CYU_KBN_HJU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_3 & "　計"
                    Case CYU_KBN_TOK
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_4 & "　計"
                    Case CYU_KBN_BOU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_E & "　計"
                End Select
                                            
                                            
                Y_SYUKA(Row, ColDen_Su) = Format(Cyu_Kei_Sumi, "#,##0") & "/" & Format(Cyu_Kei_ALL, "#,##0")
                If Cyu_Kei_ALL <> 0 Then
                    Ritu = Cyu_Kei_Sumi / Cyu_Kei_ALL
                Else
                    Ritu = 0
                End If
                Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                            
                                            
                
                Cyu_Kei_Sumi = 0
                Cyu_Kei_ALL = 0
                Save_Cyu_Kbn = Left(StrConv(Y_SYUREC.CYU_KBN, vbUnicode), 2)
                                            
            End If
                                            
                                            
                                            
            If Save_Soko_No <> Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) Then
                Row = Row + 1
                Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                            
                Select Case Save_Cyu_Kbn
                    Case CYU_KBN_TUK                '月切り
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_1
                    Case CYU_KBN_SPO                '
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_2
                    Case CYU_KBN_HJU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_3
                    Case CYU_KBN_TOK
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_4
                    Case CYU_KBN_BOU
                        Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_E
                End Select
                                        
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko_No)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Y_SYUKA(Row, ColSoko_Name) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
'2010.08.26
'                        Y_SYUKA(Row, ColSoko_Name) = "？？"
                        Y_SYUKA(Row, ColSoko_Name) = Save_Soko_No

Debug.Print "2 Save_Soko_No=" & Save_Soko_No


'2010.08.26
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                        List_Dsp_Proc = SYS_ERR
                        Exit Function
                End Select
                                                   
                Y_SYUKA(Row, ColDen_Su) = Format(Syo_Kei_Sumi, "#,##0") & "/" & Format(Syo_Kei_ALL, "#,##0")
                Ritu = Syo_Kei_Sumi / Syo_Kei_ALL
                Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                                
                Syo_Kei_Sumi = 0
                Syo_Kei_ALL = 0
                Save_Soko_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
            End If
                                            
            
            Syo_Kei_ALL = Syo_Kei_ALL + 1
            Cyu_Kei_ALL = Cyu_Kei_ALL + 1
            Sou_Kei_ALL = Sou_Kei_ALL + 1
            
            If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
                Syo_Kei_Sumi = Syo_Kei_Sumi + 1
                Cyu_Kei_Sumi = Cyu_Kei_Sumi + 1
                Sou_Kei_Sumi = Sou_Kei_Sumi + 1
            End If
        
        
        End If
        com = BtOpGetNext
    
    Loop
    
    
    If Not FAST_F Then
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                        
        Select Case Save_Cyu_Kbn
            Case CYU_KBN_TUK                '月切り
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_1
            Case CYU_KBN_SPO                '
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_2
            Case CYU_KBN_HJU
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_3
            Case CYU_KBN_TOK
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_4
            Case CYU_KBN_BOU
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_E
        End Select
                                    
        Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko_No)
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Y_SYUKA(Row, ColSoko_Name) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
            Case BtErrKeyNotFound
'2010.08.26
'                        Y_SYUKA(Row, ColSoko_Name) = "？？"
                        Y_SYUKA(Row, ColSoko_Name) = Save_Soko_No


Debug.Print "3 Save_Soko_No=" & Save_Soko_No


'2010.08.26
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                List_Dsp_Proc = SYS_ERR
                Exit Function
        End Select
                                               
        Y_SYUKA(Row, ColDen_Su) = Format(Syo_Kei_Sumi, "#,##0") & "/" & Format(Syo_Kei_ALL, "#,##0")
        If Syo_Kei_ALL <> 0 Then
            Ritu = Syo_Kei_Sumi / Syo_Kei_ALL
        Else
            Ritu = 0
        End If
        Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                        
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                        
        Select Case Save_Cyu_Kbn
            Case CYU_KBN_TUK                '月切り
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_1 & "　計"
            Case CYU_KBN_SPO                '
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_2 & "　計"
            Case CYU_KBN_HJU
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_3 & "　計"
            Case CYU_KBN_TOK
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_4 & "　計"
            Case CYU_KBN_BOU
                Y_SYUKA(Row, ColCyu_Kbn) = CYU_KBN_E & "　計"
        End Select
            
        Y_SYUKA(Row, ColDen_Su) = Format(Cyu_Kei_Sumi, "#,##0") & "/" & Format(Cyu_Kei_ALL, "#,##0")
        If Cyu_Kei_ALL <> 0 Then
            Ritu = Cyu_Kei_Sumi / Cyu_Kei_ALL
        Else
            Ritu = 0
        End If
        Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                        
                                        
        Row = Row + 1
        Y_SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                        
        Y_SYUKA(Row, ColCyu_Kbn) = "総合計"
            
        Y_SYUKA(Row, ColDen_Su) = Format(Sou_Kei_Sumi, "#,##0") & "/" & Format(Sou_Kei_ALL, "#,##0")
        If Cyu_Kei_ALL <> 0 Then
            Ritu = Sou_Kei_Sumi / Sou_Kei_ALL
        Else
            Ritu = 0
        End If
        Y_SYUKA(Row, ColProgress) = Format(Ritu * 100, "#0") & "%"
                                        
    End If

    
    Text(ptxDATE_YY).Text = Left(Format(Now, "yyyymmdd"), 4)
    Text(ptxDATE_MM).Text = Mid(Format(Now, "yyyymmdd"), 5, 2)
    Text(ptxDATE_DD).Text = Right(Format(Now, "yyyymmdd"), 2)
    Text(ptxTIME_HH).Text = Left(Format(Now, "HHmmss"), 2)
    Text(ptxTIME_MM).Text = Mid(Format(Now, "HHmmss"), 3, 2)
        
                                    'DBテーブルリンク
'    Y_SYUKA.QuickSort Min_Row, (Y_SYUKA.UpperBound(1)), 1, XORDER_ASCEND, XTYPE_STRING
    
    Set TDBGrid1.Array = Y_SYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    
        
    Call Input_UnLock
    
    List_Dsp_Proc = False
    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1060151.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060151)


    F1060151.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)

Dim sts As Integer
    
    Select Case Index
        Case 7                              '最新表示
            If List_Dsp_Proc Then           '集計＆表示
                Unload Me
            End If
            Command(7).SetFocus
        
        Case 11                             '終了
            Unload Me
    End Select
    
End Sub


Private Sub Form_Activate()
                                '集計＆表示
    If List_Dsp_Proc Then
        Unload Me
    End If
            
    Command(7).SetFocus

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
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

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
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
            F1060151.Caption = "作業監視モニター（" + RTrim(JGYOBU_T(i).NAME) + ")" + LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
                                '倉庫マスタOPEN
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データOPEN
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データOPEN
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    End Sub



Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定")
        End If
    End If
                                            '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060151 = Nothing

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
    F1060151.Caption = "作業監視モニター（" + RTrim(JGYOBU_T(Index).NAME) + ")" + LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

