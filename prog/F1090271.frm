VERSION 5.00
Begin VB.Form F1090271 
   BackColor       =   &H00C0C0C0&
   Caption         =   "在庫差異チェック処理（Active） ([F109027]2016.03.07 13:30)"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   7320
   ControlBox      =   0   'False
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label3 
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4305
      TabIndex        =   6
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2310
      TabIndex        =   5
      Top             =   3600
      Width           =   2010
   End
   Begin VB.Label Label3 
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4305
      TabIndex        =   4
      Top             =   3120
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2310
      TabIndex        =   3
      Top             =   3120
      Width           =   2010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "ホストデータ集計中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データ初期化中"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Caption         =   "在庫差異チェック処理(Active)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   6720
   End
End
Attribute VB_Name = "F1090271"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type YUKO_SOKO_TBL             '有効ﾎｽﾄ倉庫取り込みテーブル
    HS_SOKO As String * 8
    NAIGAI As String * 1
End Type

Private SOKO_T() As YUKO_SOKO_TBL        '倉庫情報

Private HS_NaiG As String               '国内外（決定内容）････　ﾎｽﾄﾃﾞｰﾀ内容により設定

Private MAX_SOKO    As Integer

Private ZENKAI_YMD      As String       '前回処理年月日
Private ZENZENKAI_YMD   As String       '前々回処理年月日

Private Function SumZ_Init() As Integer
'----------------------------------------------------------------------------
'                   「在庫集計データ」ホスト在庫クリアー処理
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer

Dim In_Cnt  As Long


    SumZ_Init = True
    
    In_Cnt = 0
    
    com = BtOpGetFirst

    Do
                
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫集計データ")
                    Exit Function
            End Select
        Loop
        
        If sts Then
            Exit Do
        End If
        
        
        
        In_Cnt = In_Cnt + 1
        Label2(0) = Format(In_Cnt, "#,##0")
        
        
        
        '前回→前々回
        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
        
        Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
        Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")

        
        
        
        
        
        
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(SUMZREC.T_Zai_Qty, vbUnicode))
        
        Do
            sts = BTRV(BtOpUpdate, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "在庫集計データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
        
        DoEvents
    Loop

    SumZ_Init = False

End Function

Private Function SumZ_Update(JGYOBU As String) As Integer
'----------------------------------------------------------------------------
'                   「在庫集計データ」ホスト在庫更新処理
'----------------------------------------------------------------------------

Dim i               As Integer
Dim j               As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim T_Zai_Qty       As Long
Dim SAI_QTY         As Long
        
Dim ans             As Integer
        
Dim In_Cnt          As Long
        
        
Dim Skip_Flg        As Boolean
        
Dim NAIGAI          As String * 1



    SumZ_Update = True

    In_Cnt = 0

    Do
        DoEvents

        Get #HS_ZAI_SSPC_No, , HS_ZAI_SSPCREC
        If Left(StrConv(HS_ZAI_SSPCREC.HS_HIN_GAI, vbUnicode), 1) < " " Then
            Exit Do
        End If







        If StrConv(HS_ZAI_SSPCREC.HS_CRLF, vbUnicode) <> vbCrLf Then



            Exit Do
        End If

        '在庫数＝０
        Skip_Flg = False
        If Not IsNumeric(StrConv(HS_ZAI_SSPCREC.HS_SURYO, vbUnicode)) Then
            Skip_Flg = True
        Else
            If CLng(StrConv(HS_ZAI_SSPCREC.HS_SURYO, vbUnicode)) = 0 Then
                Skip_Flg = True
            End If
        End If




        If Not Skip_Flg Then
            '有効データのチェック＆国内外の獲得
            Skip_Flg = True
            For i = 0 To UBound(JGYOBU_T)
                If JGYOBU = JGYOBU_T(i).CODE Then
                    For j = 0 To UBound(SOKO_T, 2)
                        If Trim(StrConv(HS_ZAI_SSPCREC.HS_SHUSI, vbUnicode)) = _
                                                        Trim(SOKO_T(i, j).HS_SOKO) Then
                            NAIGAI = SOKO_T(i, j).NAIGAI

                            Skip_Flg = False
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i





            If Not Skip_Flg Then
                '対象データ
                Call UniCode_Conv(K0_SUMZ.JGYOBU, JGYOBU)
                Call UniCode_Conv(K0_SUMZ.NAIGAI, NAIGAI)
                Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(HS_ZAI_SSPCREC.HS_HIN_GAI, vbUnicode))

                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "在庫集計データ")
                            Exit Function
                    End Select
                Loop

                If Upd_com = BtOpInsert Then
                '新規追加時、品目マスタよりホスト棚番獲得

                    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_ZAI_SSPCREC.HS_HIN_GAI, vbUnicode))
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                        'ありえないがスルー
                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                            Call UniCode_Conv(ITEMREC.ST_RETU, "")
                            Call UniCode_Conv(ITEMREC.ST_REN, "")
                            Call UniCode_Conv(ITEMREC.ST_DAN, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select


                    Call UniCode_Conv(SUMZREC.JGYOBU, JGYOBU)
                    Call UniCode_Conv(SUMZREC.NAIGAI, NAIGAI)
                    Call UniCode_Conv(SUMZREC.HIN_GAI, StrConv(HS_ZAI_SSPCREC.HS_HIN_GAI, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(SUMZREC.T_Zai_Qty, "00000000")
                    Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, "00000000")
                    Call UniCode_Conv(SUMZREC.SYK_E_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.NYUKA_YQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                    
                    
'''2007.05.17                    Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))
                    
                    
                    Call UniCode_Conv(SUMZREC.BU_ZAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, "00000000")
                    
                    
                    '2009.02.09
                    Call UniCode_Conv(SUMZREC.SAI_QTY, "00000000")
                    Call UniCode_Conv(SUMZREC.SAI_YMD, "")
                    
                    
                    
                    Call UniCode_Conv(SUMZREC.FILLER, "")
                    


                End If

                Call UniCode_Conv(SUMZREC.SUM_DT, Format(Date, "yyyymmdd"))     '2007.05.17

                Call UniCode_Conv(SUMZREC.PPSC_ZAI_QTY, Format(CLng(StrConv(SUMZREC.PPSC_ZAI_QTY, vbUnicode)) + CLng(StrConv(HS_ZAI_SSPCREC.HS_SURYO, vbUnicode)), "00000000"))

''                SAI_QTY = CLng(StrConv(SUMZREC.T_Zai_Qty, vbUnicode)) - CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
''
''                If SAI_QTY >= 0 Then
''                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "00000000"))
''                Else
''                    Call UniCode_Conv(SUMZREC.SAI_QTY, Format(SAI_QTY, "0000000"))
''                End If










                Do

                    sts = BTRV(Upd_com, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE     'これはない
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫集計データ")
                            Exit Function
                    End Select



                Loop


                In_Cnt = In_Cnt + 1
                Label2(1) = Format(In_Cnt, "#,##0")



            End If

        End If

    Loop

    SumZ_Update = False

End Function

Private Function HS_ZAI_MAIN() As Integer
'----------------------------------------------------------------------------
'                   「在庫集計データ」ホスト在庫更新処理
'----------------------------------------------------------------------------
Dim Ret         As String

Dim FileName    As String

Dim i           As Integer
Dim ans         As Integer



    HS_ZAI_MAIN = True

    For i = 0 To UBound(JGYOBU_T)
        
If JGYOBU_T(i).CODE = "4" Then
    Debug.Print
End If
    
        If HS_ZAI_SSPC_Open(1, JGYOBU_T(i).CODE) Then
            Exit For
        End If
    
    
        If SumZ_Update(JGYOBU_T(i).CODE) Then   '「在庫集計データ」ホスト在庫更新処理

            Exit Function
        End If
    
    
        Close #HS_ZAI_SSPC_No
    
    Next i

    HS_ZAI_MAIN = False

    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブが見つかりません" & FileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("ファイルが見つかりません" & FileName, vbExclamation)
        Case 76
            Beep
            ans = MsgBox("ファイルパスが見つかりません" & FileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー [HS_ZAI Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select
End Function

Private Sub Form_Activate()

Dim ans     As Integer
Dim c       As String


'起動ﾊﾟﾗﾒｰﾀ取込み
    c = StrConv(Trim(Command), vbUpperCase)

'起動確認（起動ﾊﾟﾗﾒｰﾀ="NoDialog"の時は無し）
    If c <> "/NODIALOG" Then
        If ZENKAI_YMD = Format(Now, "YYYY/MM/DD") Then
            ans = MsgBox("本日の在庫取り込み処理は終了しています。実行しますか？", vbYesNo, "確認処理")
        Else
            ans = MsgBox("「在庫差異チェック処理」実行しますか？", vbYesNo, "確認処理")
        End If

        If ans = vbNo Then
            Unload Me
        End If
    End If


    F1090271.MousePointer = vbHourglass
    Label1(1).Visible = True
    
    
    If SumZ_Init() Then             '差異データ初期化
        Unload Me
    End If
    
    Label1(1).Visible = False
    Label1(2).Visible = True
    
    
    If HS_ZAI_MAIN() Then           '差異チェック処理(メインループ)
        Unload Me
    End If
                                    'ＩＮＩ処理日付出力
    If WriteIni(App.EXEName, "ZENZENKAI_YMD", App.EXEName, ZENKAI_YMD) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & " ZENZENKAI_YMD")
        Unload Me
    End If

    If WriteIni(App.EXEName, "ZENKAI_YMD", App.EXEName, Format(Now, "YYYY/MM/DD")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & " ZENKAI_YMD")
        Unload Me
    End If



    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer
Dim j       As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                                 
    If JGYOB_TB_Set(1) Then      '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                
                                'MAX_SOKO取り込み
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Beep
        MsgBox "[MAX_SOKO]の獲得に失敗しました。処理を中止して下さい。"
        End
    Else
        If Not IsNumeric(Trim(c)) Then
            Beep
            MsgBox "[MAX_SOKO]の獲得に失敗しました。処理を中止して下さい。"
            End
        Else
            MAX_SOKO = CInt(RTrim(c))
        End If
    End If
                                
                                
                                
''    If JGYOBU_T(0).CODE = AIRCON Then
''    Else
                                    '在庫取り込み用テーブル作成
        ReDim SOKO_T(0 To UBound(JGYOBU_T), 0 To MAX_SOKO - 1)
                                    '倉庫情報取り込み
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.29
'        For i = 0 To UBound(JGYOBU_T)
'            j = 0
'            Do
'                                    '有効倉庫獲得
'                If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
'                    Beep
'                    MsgBox "倉庫情報の獲得に失敗しました。処理を中止して下さい。"
'                    End
'                End If
'
'                If Trim(c) = "**" Then  '倉庫指定終了
'                    Exit Do
'                End If
'
'
''                ReDim Preserve SOKO_T(0 To i, 0 To j)
''                ReDim Preserve SOKO_T(0 To UBound(JGYOBU_T), 0 To j)
'                SOKO_T(i, j).HS_SOKO = Trim(c)
'                                    '国内外情報獲得
'                If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
'                    Beep
'                    MsgBox "国内外情報の獲得に失敗しました。処理を中止して下さい。"
'                    End
'                End If
'
'                SOKO_T(i, j).NAIGAI = Trim(c)
'                j = j + 1
'            Loop
'
'        Next i



        For i = 0 To UBound(JGYOBU_T)
            j = 0
            Do
                                    '有効倉庫獲得
                If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                    c = "@@"
                End If

                If Trim(c) = "**" Then  '倉庫指定終了
                    Exit Do
                End If


                SOKO_T(i, j).HS_SOKO = Trim(c)
                                    '国内外情報獲得
                If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                    c = "1"
                End If

                SOKO_T(i, j).NAIGAI = Trim(c)
                j = j + 1
                If j > MAX_SOKO - 1 Then
                    Exit Do
                End If
            Loop

        Next i







'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.29
''    End If
                                '前回処理日の獲得
    If GetIni(App.EXEName, "ZENZENKAI_YMD", App.EXEName, c) Then
        ZENZENKAI_YMD = ""
    Else
        ZENZENKAI_YMD = RTrim(c)
    End If
                                
                                '前々回処理日の獲得
    If GetIni(App.EXEName, "ZENKAI_YMD", App.EXEName, c) Then
        ZENKAI_YMD = ""
    Else
        ZENKAI_YMD = RTrim(c)
    End If
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If
                                '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
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
                                            '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1090271 = Nothing

    End
End Sub


