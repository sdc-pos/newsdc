Attribute VB_Name = "PI00030com"
Option Explicit


Public pubBikou_1   As String   '備考１ 2007.07.20
Public pubBikou_2   As String   '備考２ 2007.07.20
Public pubBikou_3   As String   '備考３ 2007.07.20



'---------------------------------------------- *検索用資材注文ﾃﾞｰﾀ
'ポジショニング
Public wP_SHORDER_POS       As POSBLK
'データ・バッファ
Public wP_SHORDER_REC       As P_SHORDER_REC_Tag
'キー・データ
Public K2_wP_SHORDER        As KEY2_P_SHORDER

Public GLB_SYUSHI_F     As String           '2017.11.17

Public Function wP_SHORDER_Open(Mode As Integer) As Integer
'****************************************************
'*      「資材注文ﾃﾞｰﾀ」    ＯＰＥＮ処理
'*
'*  資材注文ﾃﾞｰﾀを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    
    wP_SHORDER_Open = True
                                    '資材注文ﾃﾞｰﾀ　フルパス取込み
    sts = GetIni("FILE", P_SHORDER_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_SHORDER]読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                ans = MsgBox("他端末でデータ使用中です。<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wP_SHORDER_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "資材注文ﾃﾞｰﾀ")
                Exit Function
        End Select
    Loop

    wP_SHORDER_Open = False

End Function

Public Function wP_SHORDER_CLOSE() As Integer

'****************************************************
'*      「資材注文ﾃﾞｰﾀ」    ＣＬＯＳＥ処理
'*
'*  資材注文ﾃﾞｰﾀを別ポインタでＣＬＯＳＥする
'*  (呼び元で終了時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'****************************************************
Dim sts As Integer
    
    wP_SHORDER_CLOSE = True
    
    sts = BTRV(BtOpClose, wP_SHORDER_POS, wP_SHORDER_REC, Len(wP_SHORDER_REC), K2_wP_SHORDER, Len(K2_wP_SHORDER), 2)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "資材注文ﾃﾞｰﾀ")
            Exit Function
    End Select

    wP_SHORDER_CLOSE = False

End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Sub Main()
    
Dim lngReturnValue      As Long
Dim strMyTitle          As String
Dim lngPrevHwnd         As Long
Dim lngTopHwnd          As Long
Dim lngThreadID1        As Long
Dim lngThreadID2        As Long
    
    
    
    GLB_SYUSHI_F = Trim(Command)


    ' 2重起動の場合は、手前に持ってきて自分自身は終了する
    strMyTitle = App.Title
    App.Title = "$" & App.Title
    lngPrevHwnd = FindWindow("ThunderRT6Main", strMyTitle)
    If lngPrevHwnd <> 0 Then
    lngTopHwnd = GetLastActivePopup(lngPrevHwnd)
    If IsIconic(lngTopHwnd) = WIN32API_TRUE Then
    lngReturnValue = ShowWindow(lngTopHwnd, SW_NORMAL)
    End If
    lngThreadID1 = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
    lngThreadID2 = GetCurrentThreadId()
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 1)
    lngReturnValue = SetForegroundWindow(lngTopHwnd)
    lngReturnValue = AttachThreadInput(lngThreadID2, lngThreadID1, 0)
    Exit Sub
    End If
    App.Title = strMyTitle




    PI000301.Show
End Sub

