VERSION 5.00
Begin VB.Form F1020551 
   BackColor       =   &H00FFFFFF&
   Caption         =   "標準棚番順入庫現品票印刷"
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
   Begin VB.ListBox List1 
      Height          =   300
      Left            =   120
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5100
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   7770
      TabIndex        =   30
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   7140
      TabIndex        =   28
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   6510
      TabIndex        =   26
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   24
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   5145
      TabIndex        =   22
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   4410
      TabIndex        =   20
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   3780
      TabIndex        =   17
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3045
      TabIndex        =   16
      Top             =   1680
      Width           =   330
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   3045
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   14
      Top             =   3360
      Width           =   4095
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
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
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7455
      TabIndex        =   29
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   6825
      TabIndex        =   27
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6195
      TabIndex        =   25
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "～"
      Height          =   255
      Index           =   4
      Left            =   5565
      TabIndex        =   23
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   4830
      TabIndex        =   21
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   4095
      TabIndex        =   19
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   3465
      TabIndex        =   18
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '透明
      Caption         =   "標準棚番"
      Height          =   255
      Index           =   0
      Left            =   1890
      TabIndex        =   15
      Top             =   1800
      Width           =   1065
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
      TabIndex        =   13
      Top             =   6480
      Width           =   180
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
      TabIndex        =   12
      Top             =   4800
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
Attribute VB_Name = "F1020551"
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

'Dim wTEXT(0 To 7)   As String
Dim wkStr_S         As String
Dim wkStr_E         As String

Dim JGYOBU_NAME     As String

Dim Printer_tbl()   As String
Dim Max_Gyo         As Integer

Private Function Input_Data_Check() As Integer

Dim i           As Integer


    Input_Data_Check = True


'    Erase wTEXT
    For i = 0 To 7
        If Trim(Text1(i).Text) = "" Then
            If i <> 0 And i <> 4 Then
                Text1(i).Text = "00"
            End If
        Else
            If IsNumeric(Text1(i).Text) Then
                If i <> 0 And i <> 4 Then
                    Text1(i).Text = Format(Val(Text1(i).Text), "00")
                End If
            Else
                Beep
                MsgBox "入力した項目はエラーです。"
                Text1(i).SetFocus
                Exit Function
            End If
        End If
    Next i


    wkStr_S = Text1(0).Text & Text1(1).Text & Text1(2).Text & Text1(3).Text
    wkStr_E = Text1(4).Text & Text1(5).Text & Text1(6).Text & Text1(7).Text

    If Len(wkStr_S) <> 8 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text1(0).SetFocus
        Exit Function
    End If

    If Len(wkStr_E) <> 8 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text1(4).SetFocus
        Exit Function
    End If

    If wkStr_S > wkStr_E Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text1(0).SetFocus
        Exit Function
    End If


    Input_Data_Check = False


End Function

Private Sub Print_Data_Select()
Dim yn              As Integer
Dim RetBuf          As String
Dim sts             As Integer
Dim com             As Integer
Dim wkList_Box      As String
Dim wk_Kbn          As String * 1
Dim wk_Bikou        As String * 15
Dim wk_Maisuu       As Integer

Dim wk_IRI_QTY      As String * 8
Dim wk_MAISU        As String * 3

Dim wkStr           As String


    List1.Clear

    Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ITEM.NAIGAI, NAIGAI_NAI$)
    Call UniCode_Conv(K6_ITEM.ST_SOKO, Trim(Text1(0).Text))
    Call UniCode_Conv(K6_ITEM.ST_RETU, Trim(Text1(1).Text))
    Call UniCode_Conv(K6_ITEM.ST_REN, Trim(Text1(2).Text))
    Call UniCode_Conv(K6_ITEM.ST_DAN, Trim(Text1(3).Text))
    Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
    com = BtOpGetGreaterEqual

    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
        Select Case sts
            Case BtNoErr
                wkStr = StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)
                If wkStr > wkStr_E Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Beep
                MsgBox "システム異常が発生しました。処理を中止して下さい。"
                Unload Me
        End Select


        wkList_Box = NAIGAI_NAI$ & " " & StrConv(ITEMREC.HIN_GAI, vbUnicode) & " "
        wkList_Box = wkList_Box & StrConv(ITEMREC.HIN_NAI, vbUnicode) & " "

        wkList_Box = wkList_Box & Space(Len(wk_IRI_QTY)) & "   "
        wkList_Box = wkList_Box & Space(Len(wk_MAISU)) & "   "
        wkList_Box = wkList_Box & Space(Len(wk_Bikou)) & "   "

        wkList_Box = wkList_Box & StrConv(ITEMREC.HIN_NAME, vbUnicode) & " "
        List1.AddItem wkList_Box

        com = BtOpGetNext
    Loop

End Sub

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

    Printer.PaperSize = vbPRPSA4
    Printer.Orientation = vbPRORPortrait   '用紙の短辺を上にして印刷
    Max_Gyo = 5


    For Gyo = 0 To UBound(Print_tbl)
        For Retu = 0 To 1
        
            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
        Next Retu
    Next Gyo

    Gyo = 0
    Retu = 0


    For wk_LOOP = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(wk_LOOP).Code = Last_JGYOBU Then
            JGYOBU_NAME = JGYOBU_T(wk_LOOP).NAME
            Exit For
        End If
    Next wk_LOOP



    For wk_LOOP = 0 To List1.ListCount - 1
        wk_Naigai = Left(List1.List(wk_LOOP), 1)
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, wk_Naigai)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Mid(List1.List(wk_LOOP), 3, 13))
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
        
'        For Maisu = 1 To CInt(Mid(List1.List(wk_LOOP), 42, 3))
            
            DoEvents
            
            If wk_Naigai = NAIGAI_NAI Then
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI1
            Else
                Print_tbl(Gyo, Retu).NAIGAI = NAIGAI2
            End If
            Print_tbl(Gyo, Retu).HIN_GAI = Mid(List1.List(wk_LOOP), 3, 13)
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
    
            Print_tbl(Gyo, Retu).IRI_QTY = Mid(List1.List(wk_LOOP), 31, 8)
            Print_tbl(Gyo, Retu).BIKOU = Mid(List1.List(wk_LOOP), 48, 15)
    
            Retu = Retu + 1
            If Retu > 1 Then
                Gyo = Gyo + 1
                If Gyo > Max_Gyo Then
                    Call Print_Sub_Proc

'2005/05/24 Upd Start #################################################################
'                    Printer.NewPage
                    Printer.EndDoc
'2005/05/24 Upd End   #################################################################

                    For Gyo = 0 To Max_Gyo
                        For Retu = 0 To 1
        
                            Print_tbl(Gyo, Retu).HIN_GAI = " "
        
                        Next Retu
                    Next Gyo

                    Gyo = 0
                End If
                Retu = 0
            End If
'        Next Maisu


    Next wk_LOOP
    
    Call Print_Sub_Proc
        
End Function
                                    '画面初期状態を設定する
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To 7
        Text1(i).Text = ""
    Next i
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
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
                                            
        Case 8                              '印刷
            If Input_Data_Check Then
                Exit Sub
            End If

'2005/05/24 Upd Start ###########################################################################
'            Beep
'            yn = MsgBox("印刷しますか？", vbYesNo + vbQuestion, "確認入力")
'            If yn = vbYes Then
'                Call Print_Data_Select

            Me.Enabled = False
            DoEvents
            Call Print_Data_Select

            Beep
            yn = MsgBox("現品票は " & Format(List1.ListCount, "0,0枚 印刷されます。") & vbCrLf & vbCrLf & _
                        "宜しいですか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
            If yn = vbYes Then
'2005/05/24 Upd Start ###########################################################################

                sts = Print_Proc()
                Printer.EndDoc
                Call Clear_Field
            End If
            
            Me.Enabled = True
            DoEvents
            
            Text1(0).SetFocus
            
        Case 11                             '終了
            Beep
            yn = MsgBox("終了しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Unload Me
            End If
            Text1(0).SetFocus
            
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
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1020551.Caption = "標準棚番順入庫現品票印刷（" + RTrim(JGYOBU_T(i).NAME) + ")"
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
                                
                                '印刷フォント設定
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020551.FontName
        .Size = F1020551.FontSize
    End With
    Set Printer.Font = NormalFont
                                
                                '画面初期設定
    
    
    Call Clear_Field
    
    
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

    Text1(0).SetFocus

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

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1020551.Caption = "標準棚番順入庫現品票印刷（" + RTrim(JGYOBU_T(Index).NAME) + "）"
    Last_JGYOBU = JGYOBU_T(Index).Code
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim RetBuf      As String
Dim i           As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn And Index < 4 Then
                If Trim(Text1(Index + 4).Text) = "" Then
                    Text1(Index + 4).Text = Text1(Index).Text
                End If
            End If

            For i = Index + 1 To 7 Step 1
                If Text1(i).Enabled Then
                    Text1(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text1(i).Enabled Then
                    Text1(i).SetFocus
                    Exit For
                End If
            Next i
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
            Printer.Print Tab(20);
            Printer.Print "*" + Print_tbl(Gyo, 1).HIN_GAI + "*"
        End If
'------------------------------------------------   2行目   ------------------
        With NormalFont
            .NAME = F1020551.FontName
            .Size = 14
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
'        Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
        Printer.Print " ";
        With NormalFont
            .NAME = F1020551.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(18);
'        Printer.Print "[" & Print_tbl(Gyo, 0).NAIGAI & "]";
        Printer.Print " ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            With NormalFont
                .NAME = F1020551.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
'            Printer.Print "[" & Trim(JGYOBU_NAME) & "]";
            Printer.Print " ";
            With NormalFont
                .NAME = F1020551.FontName
                .Size = 12
            End With
            Set Printer.Font = NormalFont
            Printer.Print Tab(67);
'            Printer.Print "[" & Print_tbl(Gyo, 1).NAIGAI & "]"
            Printer.Print " "
        End If
        Printer.Print
'------------------------------------------------   3行目   ------------------
        Printer.Print Tab(4);
'        Printer.Print "[入庫現品票]" & "          ";
'        Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text;
        Printer.Print " ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
'            Printer.Print "[入庫現品票]" & "          ";
'            Printer.Print Text(5).Text & "/" & Text(6).Text & "/" & Text(7).Text
            Printer.Print " "
        End If
'------------------------------------------------   4行目   ------------------
        With NormalFont
            .NAME = F1020551.FontName
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
            .NAME = F1020551.FontName
            .Size = 12
        End With
        Set Printer.Font = NormalFont
        Printer.Print Tab(4);
        Printer.Print "品名  ";
        Printer.Print Print_tbl(Gyo, 0).HIN_NAME;
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(53);
            Printer.Print "品名  ";
            Printer.Print Print_tbl(Gyo, 1).HIN_NAME
        End If
'------------------------------------------------   6行目   ------------------
        Printer.Print Tab(13);
'        Printer.Print "入数：";
        Printer.Print " ";
       If IsNumeric(Print_tbl(Gyo, 0).IRI_QTY) Then
            wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 0).IRI_QTY), "###0"), 5)
            wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
'            Printer.Print StrConv(wk_IRI_QTY, vbWide);
            Printer.Print " ";
        Else
            Printer.Print "　　　　　";
        End If
'        Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU;
        Printer.Print " ";
        If Len(Trim(Print_tbl(Gyo, 1).HIN_GAI)) = 0 Then
            Printer.Print
        Else
            Printer.Print Tab(62);
'            Printer.Print "入数：";
            Printer.Print " ";
            If IsNumeric(Print_tbl(Gyo, 1).IRI_QTY) Then
                wk_IRI_QTY = Right(Format(CLng(Print_tbl(Gyo, 1).IRI_QTY), "###0"), 5)
                wk_IRI_QTY = Space(Len(wk_IRI_QTY) - Len(Trim(wk_IRI_QTY))) & Trim(wk_IRI_QTY)
            
'                Printer.Print StrConv(wk_IRI_QTY, vbWide);
                Printer.Print " ";
            Else
                Printer.Print "　　　　　";
            End If
'            Printer.Print "  " & Print_tbl(Gyo, 0).BIKOU
            Printer.Print " "
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
            Printer.Print Tab(53);
            Printer.Print "標準入庫棚  ";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO & ":";
            Printer.Print Print_tbl(Gyo, 1).ST_SOKO_NAME;
            Printer.Print Tab(86);
            Printer.Print Print_tbl(Gyo, 1).ST_RETU & "-" & Print_tbl(Gyo, 1).ST_REN & "-" & Print_tbl(Gyo, 1).ST_DAN
        End If
'------------------------------------------------   7行目   ------------------
        
        If Gyo <> Max_Gyo Then
        
            With NormalFont
                .NAME = F1020551.FontName
                .Size = 14
            End With
            Set Printer.Font = NormalFont
            Printer.Print
'        With NormalFont
'            .NAME = F1020551.FontName
'            .Size = 18
'        End With
'        Set Printer.Font = NormalFont
            With NormalFont
                .NAME = F1020551.FontName
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
