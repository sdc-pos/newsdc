VERSION 5.00
Begin VB.Form F1020521 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入庫現品票印刷"
   ClientHeight    =   6930
   ClientLeft      =   2025
   ClientTop       =   2940
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
   ScaleHeight     =   6930
   ScaleWidth      =   11295
   StartUpPosition =   2  '画面の中央
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   2880
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   44
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "用紙選択"
      Height          =   975
      Left            =   360
      TabIndex        =   41
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A4"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   43
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A5"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   4320
      MaxLength       =   3
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   8280
      MaxLength       =   3
      TabIndex        =   39
      Top             =   5280
      Width           =   732
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      DragMode        =   1  '自動
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   7320
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3720
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   7320
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   10215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   7320
      MaxLength       =   40
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "12345678901234567890"
      Top             =   720
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
      Left            =   10320
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "印  刷"
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   13
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "確  定"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数合計"
      Height          =   252
      Index           =   12
      Left            =   7200
      TabIndex        =   38
      Top             =   5400
      Width           =   972
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷備考"
      Height          =   255
      Index           =   11
      Left            =   7200
      TabIndex        =   37
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "枚数"
      Height          =   255
      Index           =   10
      Left            =   6360
      TabIndex        =   36
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   255
      Index           =   9
      Left            =   5520
      TabIndex        =   35
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "入数"
      Height          =   255
      Index           =   8
      Left            =   6600
      TabIndex        =   34
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日"
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   33
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "月"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   32
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "年"
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   31
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷日付"
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   30
      Top             =   2640
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
      Left            =   120
      TabIndex        =   29
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "国内外"
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   28
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷備考"
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   26
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "印刷枚数"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（内部）"
      Height          =   255
      Index           =   14
      Left            =   2160
      TabIndex        =   24
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品番（外部）"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   23
      Top             =   840
      Width           =   1455
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1020521"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NormalFont As New StdFont           '印刷フォント
Dim Code39Font As New StdFont           '印刷フォント


Private Type Print_tbl_tag              '印刷用テーブル
    NAIGAI          As String * 2
    HIN_GAI         As String * 13
    HIN_NAI         As String * 13
    HIN_NAME        As String * 25
    IRI_QTY         As String * 8
    ST_SOKO         As String * 2
    ST_SOKO_NAME    As String * 5
    ST_RETU         As String * 2
    ST_REN          As String * 2
    ST_DAN          As String * 2
    BIKOU           As String * 15
End Type

Dim Print_tbl(0 To 6, 0 To 1) _
                    As Print_tbl_tag



Dim JGYOBU_NAME As String

Dim Printer_tbl() As String
Dim Max_Gyo     As Integer


Private Function Print_Proc() As Integer

Dim Maisu       As Integer
Dim sts         As Integer
Dim flg         As Boolean

Dim wk_LOOP     As Integer

Dim Gyo         As Integer


Dim Retu        As Integer

Dim wk_Naigai   As String * 1

Dim Wk_Printer As Printer

    Print_Proc = False

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



    For wk_LOOP = 0 To List1.ListCount - 1
        wk_Naigai = Left(List1.List(wk_LOOP), 1)
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(List1.List(wk_LOOP), 3, 20))
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
        
        For Maisu = 1 To CInt(Mid(List1.List(wk_LOOP), 49, 3))
            
            DoEvents
            
            If wk_Naigai = NAIGAI_NAI Then
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
            Else
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
            End If
            Print_tbl(Gyo, Retu).HIN_GAI = Mid(List1.List(wk_LOOP), 3, 20)
            If Not flg Then
                Print_tbl(Gyo, Retu).HIN_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                Print_tbl(Gyo, Retu).HIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Print_tbl(Gyo, Retu).ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                Print_tbl(Gyo, Retu).ST_RETU = StrConv(ITEMREC.ST_RETU, vbUnicode)
                Print_tbl(Gyo, Retu).ST_REN = StrConv(ITEMREC.ST_REN, vbUnicode)
                Print_tbl(Gyo, Retu).ST_DAN = StrConv(ITEMREC.ST_DAN, vbUnicode)
    
                Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
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
            End If
    
            Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 38, 8)
            Print_tbl(Gyo, Retu).BIKOU = Mid(List1.List(wk_LOOP), 55, 15)
    
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
        Next Maisu

    
    Next wk_LOOP
    
    Call Print_Sub_Proc
        
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

Private Sub Combo_Click(Index As Integer)
        
        Text(0).SelStart = 0
        Text(0).SelLength = Len(RTrim(Text(0).Text))
        Text(0).SetFocus
End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            Text(0).SetFocus
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select

End Sub



Private Sub Command_Click(Index As Integer)

Dim yn              As Integer
Dim RetBuf          As String
Dim sts             As Integer
Dim wkList_Box      As String
Dim wk_Kbn          As String * 1
Dim wk_Bikou        As String * 15
Dim wk_Maisuu       As Integer

Dim wk_IRI_QTY      As String * 8
Dim wk_MAISU        As String * 3



Select Case Index
        Case 0                              '確定
                                            
                                            '外部品番で読み込み
            If Len(Text(0).Text) <> 0 Then
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                If Combo(0).Text = NAIGAI1$ Then
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                    wk_Kbn = NAIGAI_NAI
                Else
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                    wk_Kbn = NAIGAI_GAI
                End If
                Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Text(1).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                        Text(2).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                    Case BtErrKeyNotFound
                        MsgBox "入力したコードは登録されていません。"
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                End Select
                                                        
            Else                            '内部品番で読み込み
                
                
                #If Center_chk Then
                    Call UniCode_Conv(K3_ITEM.JGYOBU, Last_JGYOBU)
                    If Combo(0).Text = NAIGAI1$ Then
                        Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                        wk_Kbn = NAIGAI_NAI
                    Else
                        Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                        wk_Kbn = NAIGAI_GAI
                    End If
                    Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(Text(2).Text))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                #Else
                    Call UniCode_Conv(K1_ITEM.JGYOBU, Last_JGYOBU)
                    If Combo(0).Text = NAIGAI1$ Then
                        Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_NAI$)
                        wk_Kbn = NAIGAI_NAI
                    Else
                        Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_GAI$)
                        wk_Kbn = NAIGAI_GAI
                    End If
                    Call UniCode_Conv(K1_ITEM.HIN_NAI, RTrim(Text(2).Text))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K1_ITEM, Len(K1_ITEM), 1)
                #End If
                Select Case sts
                    Case BtNoErr
                        Text(0).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                        Text(1).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        MsgBox "入力したコードは登録されていません。"
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Beep
                        MsgBox "システム異常が発生しました。処理を中止して下さい。"
                        Unload Me
                End Select
            End If
                                            'エラーチェック
            If Len(RTrim(Text(0).Text)) = 0 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(0).SetFocus
                Exit Sub
            End If
    
        
            If Len(Text(3).Text) = 0 Then
                Text(3).Text = "1"
            End If
            
            
            If Not IsNumeric(Text(3).Text) Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(3).SetFocus
                Exit Sub
            Else
                Text(3).Text = Format(CInt(Text(3).Text), "#0")
            
            End If
            If CInt(Text(3).Text) < 1 Then
                Beep
                MsgBox "入力した項目はエラーです。"
                Text(3).SetFocus
                Exit Sub
            End If
            
            If Text(8).Text = "" Then
            Else
                If Len(Trim(Text(8).Text)) = 0 Then
                Else
                    If Not IsNumeric(Text(8).Text) Then
                        Beep
                        MsgBox "入力した項目はエラーです。"
                        Text(8).SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
            Beep
            yn = MsgBox("確定しますか？", vbYesNo + vbQuestion, "確認入力")
            
            If yn = vbYes Then
                                
                wkList_Box = wk_Kbn & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode) + " "
                wkList_Box = wkList_Box & Left(StrConv(ITEMREC.HIN_NAI, vbUnicode), 13) + " "
                
                
                If Not IsNumeric(Text(8).Text) Then
                    wk_IRI_QTY = ""
                Else
                    wk_IRI_QTY = Format(CLng(Text(8).Text), "#0")
                End If
                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
                
                wkList_Box = wkList_Box & wk_IRI_QTY & "   "
                
                wk_MAISU = Format(CLng(Text(3).Text), "#0")
                wk_MAISU = Space(Len(wk_MAISU) - Len(Trim(wk_MAISU))) & Trim(wk_MAISU)
                
                wkList_Box = wkList_Box & wk_MAISU & "   "
                wk_Bikou = Text(4).Text
                wkList_Box = wkList_Box & wk_Bikou & "   "
                wkList_Box = wkList_Box & StrConv(ITEMREC.HIN_NAME, vbUnicode) + " "
                List1.AddItem wkList_Box
            End If
                        
            If Item_Update_Proc() Then
                Unload Me
            End If
            
            wk_Maisuu = CInt(Text(9).Text) - CInt(Text(10).Text) + CInt(Text(3).Text)
            
            Call Clear_Field
            Text(9).Text = Format(wk_Maisuu, "#0")
            Text(2).SetFocus
        Case 8                              '印刷
            Beep
            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                
                sts = Print_Proc()
                Printer.EndDoc
                Call Clear_Field
                List1.Clear
            End If
            
            Text(2).SetFocus
            
        Case 11                             '終了
            If List1.ListCount = 0 Then
                Unload Me
            End If
            Beep
            yn = MsgBox("終了しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Unload Me
            End If
            Text(2).SetFocus
            
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
    
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1020521.Caption = "入庫現品票印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i

                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                'デフォルト用紙サイズ取り込み
    If GetIni(App.EXEName, "DEF", "SYS", c) Then
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
        .NAME = F1020521.FontName
        .Size = F1020521.FontSize
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
    
    Combo(0).AddItem NAIGAI1$
    Combo(0).AddItem NAIGAI2$
    Combo(0).Text = NAIGAI1$
    
    Text(5).Text = Mid(Date, 1, 4)
    Text(6).Text = Mid(Date, 6, 2)
    Text(7).Text = Mid(Date, 9, 2)
    
    
    Call Clear_Field
    List1.Clear
    
    Combo1.Clear
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName <> Printer.DriverName Then
            Combo1.AddItem Pri_Name.DeviceName
        End If
    Next
    
    
    Combo1.ListIndex = 0
    
    Text(2).SetFocus
    
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
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "棚マスタ")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub

Private Sub List1_DblClick()

Dim ans     As Integer

    
    ans = MsgBox("指定行を削除しますか？", vbYesNo + vbDefaultButton1, "確認入力")
    
    If ans = vbYes Then
        List1.RemoveItem List1.ListIndex
    End If

'Dim sts As Integer
'Dim sts_QTY
'Dim Mode As Integer
'Dim wk_Index As Integer
'Dim RetBuf As String
'
'Dim wk_Naigai   As String * 1
'
'                                                'リストボックスより項目表示
'    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
'    wk_Naigai = Right(List1.List(List1.ListIndex), 1)
'    If wk_Naigai = "1" Then
'        Combo(0).ListIndex = 0
'    Else
'        Combo(0).ListIndex = 1
'    End If
'    Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
'
'    '97.10.12
'    wk_Index = List1.ListIndex
'    Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid$(List1.List(List1.ListIndex), 1, 13))
'                                                '外部品番で読み込み
'    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'    Select Case sts
'        Case BtNoErr
'            '97.10.12
'            Text(0) = Mid$(List1.List(List1.ListIndex), 1, 13)
'            Text(1) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
'            Text(2) = RTrim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
'            Text(3) = Mid$(List1.List(List1.ListIndex), 66, 3)
'            Text(10) = Mid$(List1.List(List1.ListIndex), 66, 3)
'            Text(4) = Trim(Mid$(List1.List(List1.ListIndex), 72, 10))
'            Text(8) = Trim(Mid$(List1.List(List1.ListIndex), 55, 8))
'            Text(8).SetFocus
'            List1.RemoveItem wk_Index
'
'        Case BtErrKeyNotFound           'これは無いはず
'            MsgBox "マスタ内容が変更されています。最新情報を再表示します。"
'            If Len(Text(0).Text) <> 0 Then
'                Mode = 0
'            Else
'                Mode = 1
'            End If
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'            Beep
'            MsgBox "システム異常が発生しました。処理を中止して下さい。"
'            Unload Me
'    End Select

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
    F1020521.Caption = "入庫現品票印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）"
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
                Case 0
                    If Len(Text(0).Text) <> 0 Then
                                                '外部品番で読み込み
                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                        If Combo(0).Text = NAIGAI1$ Then
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI$)
                        Else
                            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI$)
                        End If
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(Text(0).Text))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                                Text(1).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                Text(2).Text = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                
                                Text(4).Text = StrConv(ITEMREC.BIKOU, vbUnicode)
                                If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                    Text(8).Text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                Else
                                    Text(8).Text = ""
                                End If
                                Text(3).SetFocus
                  
                            Case BtErrKeyNotFound
                                MsgBox "入力したコードは登録されていません。"
                                Text(0).SetFocus

                                Exit Sub
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Beep
                                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                Unload Me
                        End Select
                    End If
                Case 2
                    If Len(Text(0).Text) = 0 Then
                        If Len(Text(2).Text) <> 0 Then
                                                '内部品番で読み込み
                            #If Center_chk Then
                                Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).Text = NAIGAI1$ Then
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                Call UniCode_Conv(K2_ITEM.HIN_NAI, RTrim(Text(2).Text))
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                            #Else
                                Call UniCode_Conv(K1_ITEM.JGYOBU, Last_JGYOBU)
                                If Combo(0).Text = NAIGAI1$ Then
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_NAI$)
                                Else
                                    Call UniCode_Conv(K1_ITEM.NAIGAI, NAIGAI_GAI$)
                                End If
                                Call UniCode_Conv(K1_ITEM.HIN_NAI, RTrim(Text(2).Text))
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K1_ITEM, Len(K1_ITEM), 1)
                            #End If
                            Select Case sts
                                Case BtNoErr
                                    Text(0).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                                    Text(1).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                    Text(4).Text = Left(StrConv(ITEMREC.BIKOU, vbUnicode), 10)
                                    If IsNumeric(StrConv(ITEMREC.IRI_QTY, vbUnicode)) Then
                                        Text(8).Text = Format(CLng(StrConv(ITEMREC.IRI_QTY, vbUnicode)), "#0")
                                    Else
                                        Text(8).Text = ""
                                    End If
                                    Text(3).SetFocus

                                Case BtErrKeyNotFound
                                    MsgBox "入力したコードは登録されていません。"
                                    Text(2).SetFocus
                    
                                    Exit Sub
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Beep
                                    MsgBox "システム異常が発生しました。処理を中止して下さい。"
                                    Unload Me
                            End Select
                        Else
                            MsgBox "入力した項目はエラーです。"
                            Exit Sub
                        End If
                    End If
            End Select
            
            If Index < 3 Then
                Text(8).SetFocus
            End If
            If Index = 8 Then
                Text(3).SetFocus
            End If
            If Index > 2 Then
                If Index < 7 Then
                    Text(Index + 1).SetFocus
                End If
            End If
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub


Private Sub Print_Sub_Proc()
                                            
Dim Gyo         As Integer
Dim wk_IRI_QTY  As String * 5
                                            
                                            
                                            
'    Printer.NewPage
                                            
    On Error GoTo Err_Proc
                                            
    For Gyo = 0 To 5
                                            
                                            
        If Len(Trim(Print_tbl(Gyo, 0).HIN_GAI)) = 0 Then
            Exit For
        End If
'------------------------------------------------   1行目   ------------------
        Set Printer.Font = Code39Font
        Printer.Print Tab(2);
        Printer.Print "*" + Print_tbl(Gyo, 0).HIN_GAI + "*";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
'2007.08.20            Printer.Print Tab(20);
            Printer.Print Tab(18);              '2007.08.20
            Printer.Print "*" + Print_tbl(Gyo, 1).HIN_GAI + "*"
        End If
'------------------------------------------------   2行目   ------------------
        With NormalFont
            .NAME = F1020521.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
        With NormalFont
            .NAME = F1020521.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(18);
        Printer.Print "[" & Print_tbl(Gyo, 0).NAIGAI & "]";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            With NormalFont
                .NAME = F1020521.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
            With NormalFont
                .NAME = F1020521.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(67);
            Printer.Print "[" & Print_tbl(Gyo, 1).NAIGAI & "]"
        End If
        Printer.Print
'------------------------------------------------   3行目   ------------------
        Printer.Print Tab(4);
        Printer.Print "[入庫現品票]" & "          ";
        Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            Printer.Print "[入庫現品票]" & "          ";
            Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F1020521.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "品番" & "  ";
        Printer.Print Print_tbl(Gyo, 0).HIN_GAI & " (";
        Printer.Print Print_tbl(Gyo, 0).HIN_NAI & ")";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(46);
            Printer.Print "品番" & "  ";
            Printer.Print Print_tbl(Gyo, 1).HIN_GAI & " (";
            Printer.Print Print_tbl(Gyo, 1).HIN_NAI & ")"
        End If
'------------------------------------------------   5行目   ------------------
        With NormalFont
            .NAME = F1020521.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "品名  ";
        Printer.Print Print_tbl(Gyo, 0).HIN_NAME;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
'            Printer.Print Tab(53);
            Printer.Print Tab(51);
            Printer.Print "品名  ";
            Printer.Print Print_tbl(Gyo, 1).HIN_NAME
        End If
'------------------------------------------------   6行目   ------------------
        Printer.Print Tab(13);
        Printer.Print "入数：";
        If IsNumeric(Print_tbl(Gyo, 0).IRI_QTY) Then
            wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 0).IRI_QTY), "###0"), 5)
            wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
            Printer.Print StrConv(wk_IRI_QTY, vbWide);
        Else
            Printer.Print "　　　　　";
        End If
        Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(62);
            Printer.Print "入数：";
            If IsNumeric(Print_tbl(Gyo, 1).IRI_QTY) Then
                wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 1).IRI_QTY), "###0"), 5)
                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
                Printer.Print StrConv(wk_IRI_QTY, vbWide);
            Else
                Printer.Print "　　　　　";
            End If
            Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU
        End If
'------------------------------------------------   6行目   ------------------
        Printer.Print Tab(4);
        Printer.Print "標準入庫棚  ";
        Printer.Print Print_tbl(Gyo, 0).ST_SOKO & ":";
        Printer.Print Print_tbl(Gyo, 0).ST_SOKO_NAME;
        Printer.Print Tab(37);
        Printer.Print Print_tbl(Gyo, 0).ST_RETU & "-" & Print_tbl(Gyo, 0).ST_REN & "-" & Print_tbl(Gyo, 0).ST_DAN;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
'            Printer.Print Tab(53);
            Printer.Print Tab(51);
            Printer.Print "標準入庫棚  ";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO & ":";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO_NAME;
            Printer.Print Tab(86);
            Printer.Print Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN
        End If
'------------------------------------------------   7行目   ------------------
        
        If Gyo <> Max_Gyo Then
        
            With NormalFont
                .NAME = F1020521.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print
'        With NormalFont
'            .NAME = F1020501.FontName
'            .Size = 18
'        End With
'        Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020521.FontName
                .Size = 18
            End With
            Set Printer.Font = NormalFont
            Printer.Print
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

Private Function Item_Update_Proc() As Integer

Dim sts         As Integer
Dim ans         As Integer
Dim wk_Naigai   As String * 1

    Item_Update_Proc = True

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    
    If Combo(0).Text = NAIGAI1 Then
        wk_Naigai = NAIGAI_NAI
    Else
        wk_Naigai = NAIGAI_GAI
    End If
    
    Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(0).Text)
    
    Do
        sts = BTRV(BtOpGetEqual + 200, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                MsgBox "他でデータ変更されています。更新処理を中止します。"
                Item_Update_Proc = False
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop

    Call UniCode_Conv(ITEMREC.BIKOU, Text(4).Text)
    
    
    If Text(8).Text = "" Then
        Call UniCode_Conv(ITEMREC.IRI_QTY, "")
    Else
        If Len(Trim(Text(8).Text)) = 0 Then
            Call UniCode_Conv(ITEMREC.IRI_QTY, "")
        Else
            Call UniCode_Conv(ITEMREC.IRI_QTY, Format(CLng(Text(8).Text), "00000000"))
        End If
    End If

    Do
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Item_Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Exit Function
        End Select
    Loop

    Item_Update_Proc = False


End Function
