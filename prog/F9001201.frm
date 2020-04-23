VERSION 5.00
Begin VB.Form F9001201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "奈良ﾚﾝｼﾞ用入庫現品票印刷"
   ClientHeight    =   3465
   ClientLeft      =   2025
   ClientTop       =   2940
   ClientWidth     =   11505
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
   ScaleHeight     =   3465
   ScaleWidth      =   11505
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Left            =   3720
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   37
      Top             =   240
      Width           =   6975
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   35
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   33
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   6510
      MaxLength       =   2
      TabIndex        =   31
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   5985
      MaxLength       =   2
      TabIndex        =   29
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   27
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   4725
      MaxLength       =   2
      TabIndex        =   25
      Top             =   840
      Width           =   405
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   23
      Top             =   840
      Width           =   405
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2565
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   21
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "用紙選択"
      Height          =   975
      Left            =   360
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   7965
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1320
      Width           =   732
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3675
      MaxLength       =   2
      TabIndex        =   0
      Top             =   840
      Width           =   405
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
      Left            =   10425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   9585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   8745
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印 刷"
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
      Left            =   7905
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   6585
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   4905
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "検 索"
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
      Left            =   4065
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   2745
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
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
      Left            =   1905
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "上 書"
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
      Left            =   1065
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "新 規"
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
      Left            =   225
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ﾌｧｲﾙ名"
      Height          =   255
      Index           =   8
      Left            =   2880
      TabIndex        =   36
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7455
      TabIndex        =   34
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6930
      TabIndex        =   32
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6405
      TabIndex        =   30
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "～"
      Height          =   255
      Index           =   3
      Left            =   5670
      TabIndex        =   28
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   5145
      TabIndex        =   26
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   4620
      TabIndex        =   24
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   0
      Left            =   4095
      TabIndex        =   22
      Top             =   960
      Width           =   150
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数合計"
      Height          =   255
      Index           =   12
      Left            =   6885
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "棚番範囲"
      Height          =   255
      Index           =   4
      Left            =   2625
      TabIndex        =   15
      Top             =   960
      Width           =   1095
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
      Left            =   225
      TabIndex        =   14
      Top             =   2640
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9001201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '印刷フォント
Dim Code39Font As New StdFont           '印刷フォント


Private Type Print_tbl_tag              '印刷用テーブル
    NAIGAI          As String * 2
    HIN_GAI         As String * 20
    HIN_NAI         As String * 13
    HIN_NAME        As String
    IRI_QTY         As String * 8
    ST_SOKO         As String * 2
    ST_SOKO_NAME    As String * 5
    ST_RETU         As String * 2
    ST_REN          As String * 2
    ST_DAN          As String * 2
    BIKOU           As String
    GENSAN          As String * 22
    SHIIRE_WORK_CENTER As _
                       String * 8
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag
 


Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer


Dim Err_Log_F   As String


Dim GENSANKOKU  As Boolean

Private Const Update_day$ = "[F900120] 2012.06.04 16:45"


Private Function Print_Proc(Tanaban As String, NAIGAI As String, HIN_GAI As String, Nyuka_DT As String, Qty As String, Gyo As Integer, Retu As Integer, BIKOU As String) As Integer

Dim Maisu       As Integer
Dim sts         As Integer
Dim flg         As Boolean





    Print_Proc = False
        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
    flg = False
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            flg = True
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
        
            
            
    If NAIGAI = NAIGAI_NAI Then
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
    Else
        Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
    End If
    Print_tbl(Gyo, Retu).HIN_GAI = HIN_GAI
    If Not flg Then
        Print_tbl(Gyo, Retu).HIN_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
        Print_tbl(Gyo, Retu).HIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Print_tbl(Gyo, Retu).ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        Print_tbl(Gyo, Retu).ST_RETU = StrConv(ITEMREC.ST_RETU, vbUnicode)
        Print_tbl(Gyo, Retu).ST_REN = StrConv(ITEMREC.ST_REN, vbUnicode)
        Print_tbl(Gyo, Retu).ST_DAN = StrConv(ITEMREC.ST_DAN, vbUnicode)

        Print_tbl(Gyo, Retu).GENSAN = StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)
        Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode)

        Call UniCode_Conv(K0_SOKO.Soko_No, Mid(Tanaban, 1, 2))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = Left(StrConv(SOKOREC.SOKO_NAME, vbUnicode), 5)
            Case BtErrKeyNotFound
                Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
            Case Else
                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                Beep
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Unload Me
        End Select
    
    
    
    
    Else
        Print_tbl(Gyo, Retu).HIN_NAI = " "
        Print_tbl(Gyo, Retu).HIN_NAME = " "
        Print_tbl(Gyo, Retu).ST_SOKO = " "
        
        Print_tbl(Gyo, Retu).ST_RETU = " "
        Print_tbl(Gyo, Retu).ST_REN = " "
        Print_tbl(Gyo, Retu).ST_DAN = " "
        Print_tbl(Gyo, Retu).ST_SOKO_NAME = " "
    
        Print_tbl(Gyo, Retu).GENSAN = ""
        Print_tbl(Gyo, Retu).SHIIRE_WORK_CENTER = ""
    
    
    End If

    Print_tbl(Gyo, Retu).IRI_QTY = Qty
    Print_tbl(Gyo, Retu).BIKOU = BIKOU


    
    
        
End Function
                                    
                                    '画面初期状態を設定する
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 4
        Text(i).Text = ""
    Next i
    Text(8).Text = ""

    Text(9).Text = "0"
    Text(10).Text = "0"
End Sub





Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim sts             As Integer
Dim i               As Integer




Select Case Index
        
        
        
        
        
        Case 0, 1
        
            If Data_Make_Proc(Index) Then
                Unload Me
            End If
        
        
        
        
        
        
        Case 4
        
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
        
            Beep
            yn = MsgBox("枚数検索しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                sts = Maisu_keisan_Proc()
                
                If sts Then
                    Unload Me
                End If
        
            End If
        
            Text(0).SetFocus
        
        
        Case 8                              '印刷
            
            
            For i = 0 To 7
                Select Case i
                    Case 0, 1, 2, 3
                    
                    
                        If Trim(Text(i).Text) = "" Then
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                    
                    Case 4, 5, 6, 7
                
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        Else
                            If IsNumeric(Text(i).Text) Then
                                Text(i).Text = Format(CInt(Text(i).Text), "00")
                            End If
                        End If
                
                
                
                End Select
            Next i
            
            
            
            
            
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                
                
                
                
                If Print_Main_Proc() Then
                    Unload Me
                End If
                
                
                Printer.EndDoc
            
            
            End If
            
            Text(0).SetFocus
            
        Case 11                             '終了
            Unload Me
            
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
Dim Pri_Name    As Printer
Dim DEF         As String
    
    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "奈良ﾚﾝｼﾞ用入庫現品票印刷処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
    
                                'ログファイル名取り込み
    If GetIni(App.EXEName, "ERR_LOG", App.EXEName, c) Then
        Err_Log_F = LOG_F
    Else
        Err_Log_F = Trim(c)
    End If
    
    
    
                                '原産国マスタ更新有無
    If GetIni(App.EXEName, "GENSANKOKU", App.EXEName, c) Then
        GENSANKOKU = False
    Else
        If Trim(c) = "1" Then
            GENSANKOKU = True
        Else
            GENSANKOKU = False
        End If
    End If
    
    
    
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        If JGYOBU_T(i).CODE = SHIZAI Then
        Else
            SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
    
            If JGYOBU_T(i).CODE = Last_JGYOBU Then
                F9001201.Caption = "奈良ﾚﾝｼﾞ用入庫現品票印刷(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
                SubMenu(i).Checked = True
                LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
            Else
                SubMenu(i).Checked = False
            End If
        End If
    Next i

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenRead) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '原産国ＯＰＥＮ
    If GENSAN_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '棚ＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                
                                '品番－棚ＯＰＥＮ
    If ITEM_LOC_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                'デフォルト用紙サイズ取り込み
    If GetIni(App.EXEName, "DEF", App.EXEName, c) Then
        c = ""
    End If
    DEF = RTrim(c)
                                
                                '印刷フォント設定
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '印刷フォント設定
    With NormalFont
        .NAME = F9001201.FontName
        .Size = F9001201.FontSize
    End With
    Set Printer.Font = NormalFont
                                
                                '画面初期設定
    
    If DEF = Trim(Option1(0).Caption) Then
        Option1(0).Value = True
        Option1(1).Value = False
    Else
        If DEF = Trim(Option1(1).Caption) Then
            Option1(0).Value = False
            Option1(1).Value = True
        Else
            Option1(0).Value = True
            Option1(1).Value = False
        End If
    End If
    
    Combo1.Clear
    For Each Pri_Name In Printers
        If Trim(Pri_Name.DeviceName) = Trim(Printer.DeviceName) Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    Combo1.ListIndex = 0
    
    
    For Each Pri_Name In Printers
        If Trim(Pri_Name.DeviceName) <> Trim(Combo1.Text) Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    
    Combo1.ListIndex = 0
    
    Text(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    
                                            '原産国マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "原産国マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    Set F9001201 = Nothing


    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F9001201.Caption = "奈良ﾚﾝｼﾞ用入庫現品票印刷(" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Update_day
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
Dim RetBuf      As String
Dim i           As Integer
Dim sts         As Integer
Dim sts_QTY     As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            Select Case Index
                Case 0, 1, 2, 3
                
                
                
                    For i = 0 To 3
                    
                        If Trim(Text(Index).Text) = "" Then
                        Else
                            Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)
                    
                        End If
                    Next

                
                
                    If Trim(Text(Index).Text) = "" Then
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
                
                Case 4, 5, 6, 7
            
                    
                    
                    For i = 4 To 7
                    
                        If Trim(Text(Index).Text) = "" Then
                        Else
                            Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)
                    
                        End If
                    Next
                   
                    If Trim(Text(Index).Text) = "" Then
                        Text(Index).Text = "zz"
                    Else
                        If IsNumeric(Text(Index).Text) Then
                            Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                        End If
                    End If
            
            
            
            End Select
        
        
            For i = Index + 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        
        
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled And Not Text(i).Locked Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        
        
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF2
            Command(1).Value = True
        Case vbKeyF5
            Command(4).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub


Private Sub Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
Dim wkGENSAN    As String * 15
                                            
    On Error GoTo Err_Proc
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If


'------------------------------------------------   1行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(20);
        Printer.Print "入庫現品票";
        Printer.Print Tab(47);
        Printer.Print Trim(JGYOBU_NAME);
        
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(80);
            Printer.Print "入庫現品票";
            Printer.Print Tab(104);
            Printer.Print Trim(JGYOBU_NAME)
        End If

'------------------------------------------------   2行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 6
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   3行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 0).HIN_GAI, 14)) + "*";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(23);
            Printer.Print "*" + Trim(Left(Print_tbl(Gyo, 1).HIN_GAI, 14)) + "*"
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   5行目   ------------------
       With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "品番";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print " " & Left(Print_tbl(Gyo, 0).HIN_GAI, 14);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "(" & Left(Print_tbl(Gyo, 0).HIN_NAI, 14) & ")";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            
            
            Printer.Print "品番";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print " " & Left(Print_tbl(Gyo, 1).HIN_GAI, 14);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "(" & Left(Print_tbl(Gyo, 1).HIN_NAI, 14) & ")"
        End If
'------------------------------------------------   6行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   7行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print "品名" & " " & LeftB(Print_tbl(Gyo, 0).HIN_NAME, 80);
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print "品名" & " " & LeftB(Print_tbl(Gyo, 1).HIN_NAME, 80)
        End If
'------------------------------------------------   8行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   9行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "　　入数" & ":";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print Format(Print_tbl(Gyo, 0).IRI_QTY, "#0");
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(30);
        Printer.Print "入荷日" & ":";
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 18
        End With
        Set Printer.Font = NormalFont
        Printer.Print " ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            
            Printer.Print "　　入数" & ":";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print Format(Print_tbl(Gyo, 1).IRI_QTY, "#0");
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(88);
            Printer.Print "入荷日" & ":";
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print " "
        End If
'------------------------------------------------   10行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 4
        End With
        Set Printer.Font = NormalFont
        Printer.Print
'------------------------------------------------   11行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 0).ST_SOKO & "-" & Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
        Printer.Print Tab(30);
        Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 0).BIKOU, 40));
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Printer.Print "標準棚番" & ":" & Print_tbl(Gyo, 1).ST_SOKO & "-" & Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN;
            Printer.Print Tab(88);
            Printer.Print "　備考" & ":" & RTrim(LeftB(Print_tbl(Gyo, 1).BIKOU, 40))
        End If
'------------------------------------------------   12行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(2);
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 10
        End With
        Set Printer.Font = NormalFont
        
        
        
        wkGENSAN = Left(Print_tbl(Gyo, 0).GENSAN, 13) & Right(Print_tbl(Gyo, 0).GENSAN, 2)
        
        
        
        Printer.Print "　原産国" & ":" & wkGENSAN;
        
        
        
        Printer.Print Tab(30);
        Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 0).SHIIRE_WORK_CENTER;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(43);
            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            
            wkGENSAN = Left(Print_tbl(Gyo, 1).GENSAN, 13) & Right(Print_tbl(Gyo, 1).GENSAN, 2)
            Printer.Print "　原産国" & ":" & wkGENSAN;
            
            
            Printer.Print Tab(88);
            Printer.Print "仕入先" & ":" & Print_tbl(Gyo, 1).SHIIRE_WORK_CENTER;
        End If




'------------------------------------------------   13行目   ------------------
        With NormalFont
            .NAME = F9001201.FontName
            .Size = 8
        End With
        Set Printer.Font = NormalFont
        
        Printer.Print
        
        If Gyo <> Max_Gyo Then


            With NormalFont
                .NAME = F9001201.FontName
                .Size = 10
            End With
            Set Printer.Font = NormalFont
            Printer.Print
            
            



            If Max_Gyo <> 2 Then
            
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                Printer.Print
            Else
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 4
                End With
                Set Printer.Font = NormalFont
                Printer.Print
                With NormalFont
                    .NAME = F9001201.FontName
                    .Size = 6
                End With
                Set Printer.Font = NormalFont
                Printer.Print
            
            
            End If

        End If
    Next Gyo

    Exit Sub

Err_Proc:

    If Err.Number = 482 Then
        MsgBox "プリンターエラーが発生しました。"
    Else
        MsgBox "実行時エラー：" & Err.Number
    End If
End Sub


Private Function Maisu_keisan_Proc() As Integer


Dim com         As Integer
Dim sts         As Integer

Dim svTanaban   As String * 8
Dim SvNaigai    As String * 1
Dim SvHin_Gai   As String * 20
Dim Maisu       As Integer



    Maisu_keisan_Proc = True

    Call Input_Lock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "印刷枚数　集計中", Me.hwnd, 0)





    Call UniCode_Conv(K2_ITEM_LOC.SOKO, Text(0).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Retu, Text(1).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Ren, Text(2).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Dan, Text(3).Text)
    
    Call UniCode_Conv(K2_ITEM_LOC.JGYOBU, "")
    Call UniCode_Conv(K2_ITEM_LOC.NAIGAI, "")
    Call UniCode_Conv(K2_ITEM_LOC.HIN_GAI, "")

    com = BtOpGetGreaterEqual

    Maisu = 0

    SvNaigai = ""


    Do
        DoEvents
    
        sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K2_ITEM_LOC, Len(K2_ITEM_LOC), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEM_LOCREC.SOKO, vbUnicode) & StrConv(ITEM_LOCREC.Retu, vbUnicode) & StrConv(ITEM_LOCREC.Ren, vbUnicode) & StrConv(ITEM_LOCREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "品番－棚番")
                Exit Function
        End Select
    
    
        Maisu = Maisu + Val(StrConv(ITEM_LOCREC.Print_SU, vbUnicode))
    
        com = BtOpGetNext
    
    
    Loop

    Text(8).Text = Format(Maisu, "#0")


    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "印刷枚数　集計終了", Me.hwnd, 0)

    Maisu_keisan_Proc = False



End Function


Private Function Print_Main_Proc() As Integer


Dim com             As Integer
Dim sts             As Integer


Dim Wk_Printer      As Printer


Dim Gyo             As Integer


Dim Retu            As Integer

Dim wk_LOOP         As Integer

Dim Tanaban         As String

Dim Fsw             As Boolean

    Print_Main_Proc = True

    Call Input_Lock

'指定帳票用プリンタ情報取得
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Combo1.Text) Then
                Set Printer = Wk_Printer
                Exit For
        End If
    Next

    If Option1(0).Value = True Then
        Printer.PaperSize = vbPRPSA5
        Printer.Orientation = vbPRORLandscape  '用紙の長辺を上にして印刷
        Max_Gyo = 2
    Else
        Printer.PaperSize = vbPRPSA4
        Printer.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷
        Max_Gyo = 5
    End If


 



    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "印刷処理　印刷中", Me.hwnd, 0)


    For Gyo = 0 To UBound(Print_tbl)
        For Retu = 0 To 1
        
            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
        Next Retu
    Next Gyo

    Gyo = 0
    Retu = 0


    For wk_LOOP = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(wk_LOOP).CODE = Last_JGYOBU Then
            JGYOBU_NAME = JGYOBU_T(wk_LOOP).NAME
            Exit For
        End If
    Next wk_LOOP


    Fsw = True

    Call UniCode_Conv(K2_ITEM_LOC.SOKO, Text(0).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Retu, Text(1).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Ren, Text(2).Text)
    Call UniCode_Conv(K2_ITEM_LOC.Dan, Text(3).Text)
    
    Call UniCode_Conv(K2_ITEM_LOC.JGYOBU, "")
    Call UniCode_Conv(K2_ITEM_LOC.NAIGAI, "")
    Call UniCode_Conv(K2_ITEM_LOC.HIN_GAI, "")

    com = BtOpGetGreaterEqual


    Do
        DoEvents
    
        sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K2_ITEM_LOC, Len(K2_ITEM_LOC), 2)
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEM_LOCREC.SOKO, vbUnicode) & StrConv(ITEM_LOCREC.Retu, vbUnicode) & StrConv(ITEM_LOCREC.Ren, vbUnicode) & StrConv(ITEM_LOCREC.Dan, vbUnicode) > _
                    Text(4).Text & Text(5).Text & Text(6).Text & Text(7).Text Then
                    Exit Do
                End If
            
                Fsw = False
            
            Case BtErrEOF
                
                
                
                
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "品番－棚番")
                Exit Function
        End Select
                
        For wk_LOOP = 1 To Val(StrConv(ITEM_LOCREC.Print_SU, vbUnicode))
            Tanaban = ""
            If Print_Proc(Tanaban, NAIGAI_NAI, StrConv(ITEM_LOCREC.HIN_GAI, vbUnicode), "", "", Gyo, Retu, StrConv(ITEM_LOCREC.BIKOU, vbUnicode)) Then
                Exit Function
            End If
                
                
            Retu = Retu + 1
            If Retu > 1 Then
                Gyo = Gyo + 1
                If Gyo > Max_Gyo Then
                    Call Print_Sub_Proc
                    Printer.NewPage
                    For Gyo = 0 To Max_Gyo
                        For Retu = 0 To 1
        
                            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
                        Next Retu
                    Next Gyo
        
                    Gyo = 0
                End If
                Retu = 0
            End If
                
                
        
        Next wk_LOOP
    
    
        com = BtOpGetNext
    
    
    Loop

    If Not Fsw Then
        Call Print_Sub_Proc
        Printer.NewPage
    End If


    Call Input_UnLock


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "印刷処理　印刷終了", Me.hwnd, 0)

    Print_Main_Proc = False



End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F9001201.MousePointer = vbHourglass

    Call Ctrl_Lock(F9001201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F9001201)


    F9001201.MousePointer = vbDefault

End Sub

Private Function Data_Make_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   「商品化予定ファイル」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    
    
Dim fileName        As String
Dim FileNo          As Long
    

Dim wkBuf           As String
Dim wkText          As Variant

Dim wkDATE          As String * 8

Dim Skip_Flg        As Integer


Dim No              As String * 8       '№
Dim HIN_GAI         As String * 20      '対外品番
Dim IRI_QTY         As String * 8       '印刷入り数
Dim BIKOU           As String * 20      '印刷備考

Dim Tanaban         As String * 8       '棚番

Dim Print_SU        As String * 8       '印刷枚数




    Data_Make_Proc = True

    Call Input_Lock

    FileNo = FreeFile
    fileName = Trim(Text1.Text)
    On Error GoTo Error_Proc

    Open fileName For Input As #FileNo

    On Error GoTo 0

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現品票ファイル　登録処理開始！！", Me.hwnd, 0)

                                    'テーブルリセット
    If Mode = 0 Then
        com = BtOpGetFirst
        
        Do
            DoEvents
        
            sts = BTRV(com, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "品目－棚マスタ")
                    Call Input_UnLock
                    Exit Function
            End Select
        
            sts = BTRV(BtOpDelete, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpDelete, "品目－棚マスタ")
                    Call Input_UnLock
                    Exit Function
            End Select
        
        
            com = BtOpGetNext
        Loop
    
    
    
        If GENSANKOKU Then
    
            com = BtOpGetFirst
            
            Do
                DoEvents
            
                sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, com, "原産国マスタ")
                        Call Input_UnLock
                        Exit Function
                End Select
            
                sts = BTRV(BtOpDelete, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, BtOpDelete, "原産国マスタ")
                        Call Input_UnLock
                        Exit Function
                End Select
            
            
                com = BtOpGetNext
            Loop
        
        End If
    
    
        com = BtOpGetFirst
        
        Do
            DoEvents
        
            sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "品目マスタ")
                    Call Input_UnLock
                    Exit Function
            End Select
        
            Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
            Call UniCode_Conv(ITEMREC.ST_RETU, "**")
            Call UniCode_Conv(ITEMREC.ST_REN, "**")
            Call UniCode_Conv(ITEMREC.ST_DAN, "**")
        
            Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
        
        
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpUpdate, "品目マスタ")
                    Call Input_UnLock
                    Exit Function
            End Select
        
        
            com = BtOpGetNext
        Loop
    
    
    
    End If


    Do Until EOF(FileNo)
        
        
        DoEvents
        
        Line Input #FileNo, wkBuf
    
    
    
    
        wkText = Split(wkBuf, vbTab, -1)
    
    
    
    
        Skip_Flg = False
    
        No = ""                         '№
        HIN_GAI = ""                    '対外品番
        IRI_QTY = ""                    '印刷入り数
        BIKOU = ""                      '印刷備考

        Tanaban = ""                    '棚番

        Print_SU = ""                   '印刷枚数
    
    
    
    
        If UBound(wkText) < 0 Then
            Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
            Skip_Flg = True
        End If
    
    
        Select Case UBound(wkText)
            Case 0
                No = wkText(0)
            Case 1
                No = wkText(0)
                HIN_GAI = wkText(1)
            Case 2
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
            Case 3
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
            Case 4
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
                Tanaban = wkText(4)
                Print_SU = "1"
            Case Else
                No = wkText(0)
                HIN_GAI = wkText(1)
                IRI_QTY = wkText(2)
                BIKOU = wkText(3)
                Tanaban = wkText(4)
                Print_SU = wkText(5)
                If Not IsNumeric(Print_SU) Then
                    Print_SU = "0"
                End If
        End Select
    
        If UBound(wkText) < 4 Then
            If Not Skip_Flg Then
                Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
            End If
            Skip_Flg = True
        End If
    
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
        Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                If Not Skip_Flg Then
                    Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
                End If
                Skip_Flg = True
            Case Else
                   Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                   Exit Function
        End Select
            
                    
                    
                    
                    
                    
                    
        Call UniCode_Conv(K0_TANA.Soko_No, Mid(Tanaban, 1, 2))
        Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
        Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
        Call UniCode_Conv(K0_TANA.Dan, Mid(Tanaban, 7, 2))
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                If Not Skip_Flg Then
                    Call Err_LOG_Proc(No, HIN_GAI, IRI_QTY, BIKOU, Tanaban, Print_SU)
                End If
                Skip_Flg = True
            Case Else
                   Call File_Error(sts, BtOpGetEqual, "棚マスタ")
                   Exit Function
        End Select
            
            
        If Not Skip_Flg Then
            Call UniCode_Conv(ITEM_LOCREC.No, No)
            Call UniCode_Conv(ITEM_LOCREC.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(ITEM_LOCREC.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(ITEM_LOCREC.HIN_GAI, HIN_GAI)
            Call UniCode_Conv(ITEM_LOCREC.IRI_QTY, IRI_QTY)
            Call UniCode_Conv(ITEM_LOCREC.BIKOU, BIKOU)
        
        
            Call UniCode_Conv(ITEM_LOCREC.SOKO, Mid(Tanaban, 1, 2))
            Call UniCode_Conv(ITEM_LOCREC.Retu, Mid(Tanaban, 3, 2))
            Call UniCode_Conv(ITEM_LOCREC.Ren, Mid(Tanaban, 5, 2))
            Call UniCode_Conv(ITEM_LOCREC.Dan, Mid(Tanaban, 7, 2))
        
        
            Call UniCode_Conv(ITEM_LOCREC.Print_SU, Format(Val(Print_SU), "00000000"))
            Call UniCode_Conv(ITEM_LOCREC.FILLER, "")
        
            sts = BTRV(BtOpInsert, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrDuplicates
                    sts = BTRV(BtOpUpdate, ITEM_LOC_POS, ITEM_LOCREC, Len(ITEM_LOCREC), K0_ITEM_LOC, Len(K0_ITEM_LOC), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpUpdate, "品番－棚マスタ")
                            Exit Function
                    End Select
                Case Else
                       Call File_Error(sts, BtOpInsert, "品番－棚マスタ")
                       Exit Function
            End Select
        
        
            If StrConv(ITEMREC.ST_SOKO, vbUnicode) = "**" Then
                Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(Tanaban, 1, 2))
                Call UniCode_Conv(ITEMREC.ST_RETU, Mid(Tanaban, 3, 2))
                Call UniCode_Conv(ITEMREC.ST_REN, Mid(Tanaban, 5, 2))
                Call UniCode_Conv(ITEMREC.ST_DAN, Mid(Tanaban, 7, 2))
        
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
        
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case Else
                           Call File_Error(sts, BtOpUpdate, "品目マスタ")
                           Exit Function
                End Select
        
            End If
        
            If GENSANKOKU Then
        
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> "" Then
            
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, NAIGAI_NAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))
                    
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(GENSANREC.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(GENSANREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(GENSANREC.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                            Call UniCode_Conv(GENSANREC.FILLER, "")
                            
                            Call UniCode_Conv(GENSANREC.INS_TANTO, App.EXEName)
                            Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                        
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                        
                        
                        
                            sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                            Select Case sts
                                Case BtNoErr
                                Case Else
                                       Call File_Error(sts, BtOpInsert, "原産国マスタ")
                                       Exit Function
                            End Select
                        
                        
                        
                        Case Else
                               Call File_Error(sts, BtOpGetEqual, "原産国マスタ")
                               Exit Function
                    End Select
                End If
            End If
        
        End If



    Loop




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "現品票ファイル　登録処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    Data_Make_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            Data_Make_Proc = False      '





        Case Else
    End Select
    Call Input_UnLock

End Function


Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Text = Trim(Data.Files(1))

End Sub

Public Sub Err_LOG_Proc(No As String, HIN_GAI As String, IRI_QTY As String, BIKOU As String, Tanaban As String, Print_SU As String)


    Call LOG_OUT(Err_Log_F, No & "," & HIN_GAI & "," & IRI_QTY & "," & BIKOU & "," & Tanaban & "," & Print_SU)



End Sub
