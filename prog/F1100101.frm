VERSION 5.00
Object = "{D4A17F03-6EDB-11D2-A6E0-0040262B3978}#2.2#0"; "CtrsWsk.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form F1100101 
   Caption         =   "スキャナ制御"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer tmrDate 
      Interval        =   1000
      Left            =   600
      Top             =   420
   End
   Begin MSWinsockLib.Winsock tcpHost 
      Index           =   0
      Left            =   120
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "設定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   360
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3720
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   3120
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox errText 
      Alignment       =   2  '中央揃え
      BackColor       =   &H000000FF&
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
      Left            =   90
      TabIndex        =   7
      Text            =   "LOG.TXTを確認"
      Top             =   4920
      Width           =   6135
   End
   Begin VB.TextBox errText 
      Alignment       =   2  '中央揃え
      BackColor       =   &H000000FF&
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
      Left            =   90
      TabIndex        =   6
      Top             =   4320
      Width           =   6180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6000
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   45
      TabIndex        =   4
      Top             =   3180
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   2760
      Width           =   8055
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   45
      TabIndex        =   2
      Top             =   3600
      Width           =   8055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "業務終了"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7680
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "業務開始"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6960
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Line4 
      X1              =   5760
      X2              =   5760
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   8280
      X2              =   8280
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   8280
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   8280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblINI 
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   25
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblINI 
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   24
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label lblINI 
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   23
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblINI 
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   22
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "監視モニタ"
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
      Index           =   4
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "2013/06/12 10：99：99"
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
      Left            =   3120
      TabIndex        =   20
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "　　起動日時："
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
      Left            =   1200
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3420
      TabIndex        =   16
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "自動起動時刻："
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
      Index           =   2
      Left            =   1200
      TabIndex        =   14
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'ﾌﾗｯﾄ
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3540
      TabIndex        =   11
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "自動終了時刻："
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
      Left            =   1200
      TabIndex        =   9
      Top             =   1860
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8460
   End
   Begin CTRSWSKLib.CtrsWsk CtrsWsk1 
      Left            =   270
      Top             =   1860
      _Version        =   131074
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "F1100101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

'▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
'送信済みデータ管理バッファ
Private gbl_RespBuf()   As String
Private gbl_SockConnect As Integer
Private gbl_RecvBuf()   As String
Private gbl_RecvIndex() As Integer
Private gbl_RecvCnt     As Integer
Private gbl_RecvFlg     As Boolean
'▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加


Dim Auto_Off    As Integer      '2013.06.06



'Private MAIN_TITLE      As String   '2014.07.01            2017.12.09 移動


'***************************** ＦＩＬＥ　ＤＥＬＥＴＥ確認忘れるな！！

    


'''''''''''''''''''''''''''''''' Label_File_Make_ProcのﾁｪｯｸＯＫ？　***********************************************
'''''''''''''''''''''''''''''''' 20110.11.30のﾁｪｯｸＯＫ？　***********************************************

'[2014/02/10 - M.MATSUYAMA 変更(Ver2.0.0)]
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.03.28 13:15) "
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.10.28 15:30) 広島事 品番バーコード空白対応"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.10.29 16:30) 小野 産機検品 品番読替対応"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.10.30 13:30) 大阪事 集合梱包品番追加"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.12.12 12:00) 小野 生産計画要因 海外供給区分追加"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.12.24 14:00) 奈良 商品化完了登録在庫計上"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.12.25 11:00) 商品化完了登録在庫計上時 移動履歴に入庫数を表示"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2019.12.27 15:00) IPアドレスが同一でも受信エラーにならない様に修正"
'Private Const LAST_UPDATE_DAY$ = "([F110010] 2020.03.24 16:30) 荷積み明細対応"
Private Const LAST_UPDATE_DAY$ = "([F110010] 2020.04.03 15:00) 受信サイズオーバー エラーメッセージ変更"


Private Sub CtrsWsk1_OnSendFile(ByVal intID As Integer, ByVal strFileName As String)

    Call WriteLogMsg("ファイルを送信しました。送信ファイル名（" & strFileName & "）", FNC_FILESEND, intID, , icoMessage)

End Sub

'▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
'*******************************************************************************


' 日付更新監視タイマ
' process   :   アプリケーションが開始されてから日付が更新されたかどうかを監視する
' input     :   なし
' output    :   なし
' return    :   なし
'*******************************************************************************
Private Sub tmrDate_Timer()
    '----- 日付更新の確認 -----
    If DateDiff("d", gbl_StartApp, Now) <> 0 Then
        '----- 日付が変わっている場合 -----
        
        Call WriteLogMsg("日付が変わったためログファイル名を更新します。", FNC_DATEMONITOR, , , icoMessage)
        gbl_StartApp = Now
        
        '----- アプリケーションログファイル名を更新 -----
        gbl_LogCfg.m_LogFName = GetFullPath(gbl_LogCfg.m_LogPath, App.EXEName) & "_" & Format$(gbl_StartApp, "yyyymmdd") & ".log"
        
        '----- ログファイルの保存期間チェック -----
        If gbl_LogCfg.m_LogSave > 0 Then
            '----- アプリケーションログファイルチェック -----
            Call DeleteLogFile(App.EXEName, gbl_LogCfg.m_LogSave)
        End If
    End If
End Sub

'*******************************************************************************
' Winsock コントロール オープン
' process   :   Winsock コントロールがオープンされた際の処理をおこなう
' input     :   Index       インデックス番号
' output    :   なし
' return    :   なし
'*******************************************************************************
Private Sub tcpHost_Connect(Index As Integer)
    Dim strName As String
    
    strName = IIf(Len(tcpHost(Index).RemoteHost) = 0, _
                    tcpHost(Index).RemoteHostIP, _
                    tcpHost(Index).RemoteHost & " (" & tcpHost(Index).RemoteHostIP & ")")
    
    Call WriteLogMsg("クライアント(" & strName & ")との接続しました。", FNC_SOCKCONNECT, , strName, icoMessage)
End Sub

'*******************************************************************************
' Winsock コントロール クローズ
' process   :   Winsock コントロールがクローズされた際の処理をおこなう
' input     :   Index       インデックス番号
' output    :   なし
' return    :   なし
'*******************************************************************************
Private Sub tcpHost_Close(Index As Integer)
    Dim strName As String
    
    '----- クローズ処理 -----
    tcpHost(Index).Close
    
    '----- 変数の初期化 -----
    gbl_RespBuf(Index) = ""
    
'[2014/04/07 - M.MATSUYAMA 削除(Ver2.0.5)] 下記処理を削除
    '----- 受信キューバッファを初期化 -----
'    ReDim gbl_RecvBuf(0) As String
'    ReDim gbl_RecvIndex(0) As Integer
'    gbl_RecvCnt = 0
'    gbl_RecvFlg = False
    
    strName = IIf(Len(tcpHost(Index).RemoteHost) = 0, _
                    tcpHost(Index).RemoteHostIP, _
                    tcpHost(Index).RemoteHost & " (" & tcpHost(Index).RemoteHostIP & ")")
    
    Call WriteLogMsg("クライアント(" & strName & ")との接続を閉じました。", FNC_SOCKCLOSE, , strName, icoMessage)
End Sub

'*******************************************************************************
' Winsock コントロール 接続要求
' process   :   Winsock コントロールで接続要求を受け取った際の処理をおこなう
' input     :   Index       インデックス番号
'           :   requestID   接続要求識別子
' output    :   なし
' return    :   なし
'*******************************************************************************
Private Sub tcpHost_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim intLoop As Integer
    Dim intConn As Integer
    Dim strName As String
    
    If Index = 0 Then
        '----- 同一ＩＰのコントロールがあればクローズする
        For intLoop = 1 To tcpHost.UBound
            If tcpHost(0).RemoteHostIP = tcpHost(intLoop).RemoteHostIP Then
'                tcpHost(intLoop).Close     2019/12/26 IPアドレスが同一でも受信エラーにならないように修正
                
                strName = IIf(Len(tcpHost(intLoop).RemoteHost) = 0, _
                                tcpHost(intLoop).RemoteHostIP, _
                                tcpHost(intLoop).RemoteHost & " (" & tcpHost(intLoop).RemoteHostIP & ")")
        
                Call WriteLogMsg("同一IP(" & strName & ")のコントロールがあった為、クローズしました。(接続番号 : " & CStr(intLoop) & ")", FNC_SOCKCONNREQ, , strName, icoMessage)
                
                Exit For
            End If
        Next intLoop
        
        '----- 未接続のコントロールを検索 -----
        intConn = -1
        For intLoop = 1 To tcpHost.UBound
            If tcpHost(intLoop).State = sckClosed Then
                intConn = intLoop
                Exit For
            End If
        Next intLoop
        
        '----- クライアント用コントロールを初期化 -----
        If intConn < 0 Then
            '----- 新規接続用にコントロールを追加 -----
            gbl_SockConnect = tcpHost.UBound + 1
            
            '----- Winsock コントロール追加 -----
            Call Load(tcpHost(gbl_SockConnect))
            
            '----- グローバル変数を拡張 -----
            ReDim Preserve gbl_RespBuf(gbl_SockConnect) As String
            gbl_RespBuf(gbl_SockConnect) = ""
        
            intConn = gbl_SockConnect
        End If
        
        '----- 接続要求を処理 -----
        tcpHost(intConn).LocalPort = gbl_SockCfg.m_LocalPort
        Call tcpHost(intConn).Accept(requestID)
        
'[2014/03/05 - M.MATSUYAMA 削除(Ver2.0.3)] DoEventsを削除
        'DoEvents
        
        strName = IIf(Len(tcpHost(intConn).RemoteHost) = 0, _
                        tcpHost(intConn).RemoteHostIP, _
                        tcpHost(intConn).RemoteHost & " (" & tcpHost(intConn).RemoteHostIP & ")")
        
        Call WriteLogMsg("クライアント(" & strName & ")からの接続要求に応じました。(接続番号 : " & CStr(intConn) & ")", FNC_SOCKCONNREQ, , strName, icoMessage)
    End If
End Sub

'*******************************************************************************
' Winsock コントロール データ受信
' process   :   Winsock コントロールでデータを受信した際の処理をおこなう
' input     :   Index       インデックス番号
'           :   bytesTotal  受信バイト数
' output    :   なし
' return    :   なし
'*******************************************************************************
Private Sub tcpHost_DataArrival(intIndex As Integer, ByVal bytesTotal As Long)
    Dim bytRxArray() As Byte
    Dim strRxData   As String
    Dim strTID      As String
    Dim intID       As Integer
    Dim strPID      As String
    Dim strData     As String
    Dim strName     As String
    
    On Error GoTo tcpHost_DataArrival_Error
    
    '----- 送られてきたデータをソケットから取得する -----
    Call tcpHost(intIndex).GetData(bytRxArray, vbByte, bytesTotal)
    strRxData = StrConv(bytRxArray, vbUnicode)
    
    If gbl_RecvCnt > 0 Then
        '----- 受信キューバッファを拡張 -----
        ReDim Preserve gbl_RecvBuf(gbl_RecvCnt) As String
        ReDim Preserve gbl_RecvIndex(gbl_RecvCnt) As Integer
    Else
        '----- 受信キューバッファを初期化 -----
        ReDim gbl_RecvBuf(0) As String
        ReDim gbl_RecvIndex(0) As Integer
    End If
    
    '----- 受信キューバッファにデータを追加 -----
    gbl_RecvBuf(gbl_RecvCnt) = strRxData
    gbl_RecvIndex(gbl_RecvCnt) = intIndex
    gbl_RecvCnt = gbl_RecvCnt + 1
    
    If gbl_RecvFlg = True Then
        '----- 端末IDの確認 -----
        If Len(strRxData) > 5 And IsNumeric(Mid(strRxData, 3, 3)) = True Then
            '---------------------------------------------------------
            '   メッセージが5桁以上で、3～5桁が数字として認識できる
            '   場合は端末IDとして扱う
            '---------------------------------------------------------
            intID = CInt(Mid(strRxData, 3, 3))
        Else
            intID = -1
        End If
        
        strName = IIf(Len(tcpHost(intIndex).RemoteHost) = 0, _
                        tcpHost(intIndex).RemoteHostIP, _
                        tcpHost(intIndex).RemoteHost & " (" & tcpHost(intIndex).RemoteHostIP & ")")
    
        Call WriteLogMsg("現在受信処理中の為、処理を中断します。", FNC_SOCKRECEIVE, intID, strName, icoMessage)
        Exit Sub
    End If
    
'[2014/03/17 - M.MATSUYAMA 削除(Ver2.0.4)] On Errorの位置をループ内に移動
'   On Error GoTo tcpHost_DataArrival_Error2
    
    Do
'[2014/03/17 - M.MATSUYAMA 追加(Ver2.0.4)] On Errorの位置をループ内に移動
        On Error GoTo tcpHost_DataArrival_Error2
    
        '----- 古い受信キューから処理 -----
        strRxData = gbl_RecvBuf(0)
        intIndex = gbl_RecvIndex(0)
        gbl_RecvFlg = True
    
        '----- 受信データからシーケンスNoを取得する -----
        strTID = Mid(strRxData, 1, 1)
        
        '----- メッセージタイプを取得 -----
        strPID = Mid(strRxData, 2, 1)
        
        '----- 端末IDの確認 -----
        If Len(strRxData) > 5 And IsNumeric(Mid(strRxData, 3, 3)) = True Then
            '---------------------------------------------------------
            '   メッセージが5桁以上で、3～5桁が数字として認識できる
            '   場合は端末IDとして扱う
            '---------------------------------------------------------
            intID = CInt(Mid(strRxData, 3, 3))
        Else
            intID = -1
        End If
        
        strName = IIf(Len(tcpHost(intIndex).RemoteHost) = 0, _
                        tcpHost(intIndex).RemoteHostIP, _
                        tcpHost(intIndex).RemoteHost & " (" & tcpHost(intIndex).RemoteHostIP & ")")
        
        Call WriteLogMsg("[" & ConvBinaryMsg(strRxData) & "]", FNC_SOCKRECEIVE, intID, strName, icoDownload)
        
        '----- メッセージ内容を取得 -----
        If intID < 0 Then
            strData = Mid(strRxData, 3)
        Else
            strData = Mid(strRxData, 6)
        End If
        
        If StrComp(strTID, Left$(gbl_RespBuf(intIndex), 1), vbTextCompare) = 0 _
                And StrComp(strTID, "0", vbTextCompare) <> 0 And StrComp(strTID, " ", vbTextCompare) <> 0 Then
            '----------------------------------------
            '   シーケンスIDが同じ場合は処理せずに
            '   前回の送信データをそのまま返す
            '----------------------------------------
            Call WriteLogMsg("前回の受信データと同じシーケンスNoのため同一電文を返信します。", FNC_SOCKRECEIVE, intID, strName, icoMessage)
            Call tcpHost(intIndex).SendData(gbl_RespBuf(intIndex))
        Else
            If strPID = "R" Then
                Call OnDataReceive(strPID, strTID, intID, intIndex, strData)
            End If
        End If
'        gbl_RecvFlg = False       '2018.04.13 M.Yoshizawa
        
tcpHost_DataArrival_Error2:
        
        gbl_RecvFlg = False        '2018.04.13 M.Yoshizawa
        
        
        If Err.Number <> 0 Then
            '----- 実行時エラーが発生した場合 -----
            Call WriteLogErr(Err, FNC_SOCKRECEIVE & "2")
'[2014/03/17 - M.MATSUYAMA 追加(Ver2.0.4)] エラーメッセージ追加
            Call WriteLogMsg("gbl_RecvCnt[" & CStr(gbl_RecvCnt) & "] intIndex[" & CStr(intIndex) & "]", FNC_SOCKRECEIVE, intID, strName, icoMessage)
            '----- ステータス行にエラーを表示 -----
            Text1(2).Text = Err.Description
            '----- エラーをクリアする -----
            Err.Clear
        End If
        
        '----- 処理済みのデータをクリア -----
        Dim i As Integer
        For i = 1 To gbl_RecvCnt - 1
            gbl_RecvBuf(i - 1) = gbl_RecvBuf(i)
            gbl_RecvIndex(i - 1) = gbl_RecvIndex(i)
        Next
        gbl_RecvCnt = gbl_RecvCnt - 1
        If gbl_RecvCnt > 0 Then
            '----- 受信キューが残っている場合 -----
            ReDim Preserve gbl_RecvBuf(gbl_RecvCnt - 1) As String
            ReDim Preserve gbl_RecvIndex(gbl_RecvCnt - 1) As Integer
        Else
            '----- 受信キューが残っていない場合 -----
            ReDim gbl_RecvBuf(gbl_RecvCnt) As String
            ReDim gbl_RecvIndex(gbl_RecvCnt) As Integer
        End If
        
    Loop While gbl_RecvCnt > 0

tcpHost_DataArrival_Error:
    If Err.Number <> 0 Then
        '----- 実行時エラーが発生した場合 -----
        Call WriteLogErr(Err, FNC_SOCKRECEIVE)
'[2014/03/17 - M.MATSUYAMA 追加(Ver2.0.4)] エラーメッセージ追加
        Call WriteLogMsg("gbl_RecvCnt[" & CStr(gbl_RecvCnt) & "] intIndex[" & CStr(intIndex) & "]", FNC_SOCKRECEIVE, intID, strName, icoMessage)
        '----- ステータス行にエラーを表示 -----
        Text1(2).Text = Err.Description
    End If
        
End Sub

'*******************************************************************************
' Winsock コントロール エラー
' process   :   Winsock コントロールでエラーが発生した際の処理をおこなう
' input     :   Index           インデックス番号
'           :   Number          エラー番号
'           :   Description     エラー内容
'           :   Scode           ソケット番号
'           :   Source          エラー発生元
'           :   HelpFile        ヘルプファイル名
'           :   HelpContext     ヘルプコンテキスト
'           :   CancelDisplay   表示キャンセルフラグ
' output    :   CancelDisplay   表示キャンセルフラグ(True:非表示, False:表示)
' return    :   なし
'*******************************************************************************
Private Sub tcpHost_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim strName As String
    
    '----- メッセージボックスは表示しない -----
    CancelDisplay = True
    
    If Index = 0 Then
        '----- 待ち受け状態を解除 -----
        Call tcpHost(Index).Close
    Else
        '----- クライアントとの接続時エラーの場合 -----
        Call tcpHost(Index).Close
    End If
    
    strName = IIf(Len(tcpHost(Index).RemoteHost) = 0, _
                    tcpHost(Index).RemoteHostIP, _
                    tcpHost(Index).RemoteHost & " (" & tcpHost(Index).RemoteHostIP & ")")
    
    Call WriteLogMsg("通信エラーが発生しました(" & CStr(Number) & "):" & Description, FNC_SOCKERROR, , strName, icoError)
End Sub

Private Sub CtrsWsk1_OnReceive(ByVal intID As Integer, ByVal strMsg As String, ByVal blnResp As Boolean)
    
    Call WriteLogMsg("[" & ConvBinaryMsg(strMsg) & "]", FNC_RECVDATA, intID, , icoDownload)
    
    Call OnDataReceive(IIf(blnResp = True, "R", "N"), "", intID, -1, strMsg)

End Sub
'▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加

Private Sub Command1_Click(Index As Integer)
'-------------------------------------------------------
'
'   『業務開始指示』
'       １． ポートの獲得
'       ２． ポートの開放
'-------------------------------------------------------
    
Dim ans As Integer
    
    On Error GoTo Error
    
    Select Case Index
        
        Case 0                              '業務開始
            
            '▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
            If gbl_SockCfg.m_IsListen = False Then
                '----- Winsockコントロール初期化 -----
                With tcpHost(0)
                    .Protocol = sckTCPProtocol
                    .LocalPort = gbl_SockCfg.m_LocalPort
                    .Listen
                End With
                If Err.Number = 0 Then
                    '----- 初期化に成功した場合 -----
                    Call WriteLogMsg("ネットワークリソースを初期化しました。(ローカルポート:" & CStr(gbl_SockCfg.m_LocalPort) & ")", FNC_PARENTCONN, , , icoMessage)
                    '----- フラグを更新 -----
                    gbl_SockCfg.m_IsListen = True
                Else
                    '----- 初期化に失敗した場合 -----
                    Call WriteLogErr(Err, FNC_PARENTCONN)
                End If
                '----- 受信キューバッファを初期化 -----
                ReDim gbl_RecvBuf(0) As String
                ReDim gbl_RecvIndex(0) As Integer
                gbl_RecvCnt = 0
                gbl_RecvFlg = False
            End If
            '▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
            
            CtrsWsk1.Bind LocalPort, RemotePort
            F1100101.Caption = "スキャナ制御" & " " & LAST_UPDATE_DAY
    
    
            Label1.Caption = MAIN_TITLE & "[実行中]"
    
    
            Command1(0).Enabled = False
            
            
            'Command1(0).Caption = Format(Now, "YYYY/MM/DD HH:MM:SS")       '2013.06.06
            Label4(1).Caption = Format(Now, "YYYY/MM/DD HH:MM:SS")          '2013.06.06
            
            
            Command1(1).Enabled = True
            Command1(2).Enabled = True
    
        Case 1                              '業務終了
            
            
            ans = MsgBox("本日の業務終了しますか？", vbYesNo + vbDefaultButton2, "業務終了")
            If ans = vbNo Then
                Exit Sub
            End If
            
            CtrsWsk1.Unbind
            
            '▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
            If gbl_SockCfg.m_IsListen = True Then
                Dim intIdx As Integer
                '----- Winsockコントロールクローズ -----
                For intIdx = tcpHost.UBound To tcpHost.LBound Step -1
                    If tcpHost(intIdx).State <> sckClosed Then
                        '----- 通信クローズ -----
                        Call tcpHost(intIdx).Close
                    End If
                    If Index > 0 Then
                        '----- ロード済みコントロールを破棄 -----
                        Call Unload(tcpHost(intIdx))
                        gbl_SockConnect = gbl_SockConnect - 1
                        ReDim Preserve gbl_RespBuf(gbl_SockConnect) As String
                    End If
                Next
                If Err.Number = 0 Then
                    '----- 終了処理に成功した場合 -----
                    Call WriteLogMsg("アクセスポイントのネットワークリソースを開放しました。", FNC_PARENTDISCONN, , , icoMessage)
                    '----- フラグを更新 -----
                    gbl_SockCfg.m_IsListen = False
                Else
                    '----- 終了処理に失敗した場合 -----
                    Call WriteLogErr(Err, FNC_PARENTDISCONN)
                End If
            End If
            '▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
            
            Normal_End = False              '正常終了
            
            Next_Step = 1                   '次処理起動する
            Unload Me
    
        Case 2
            'CtrsWsk1.Unbind                                    2013.06.06
                                    
            'Normal_End = False              '正常終了          2013.06.06
            'Next_Step = 0                   '次処理起動しない  2013.06.06
            Unload Me
    
    End Select
    
    Exit Sub

Error:
    MsgBox "Winsock Error= " & Err.Description    'ステータス行にエラーを表示します。
    
    Call LOG_OUT(LOG_F, "Winsock Error= " & Err.Description)
    
    Normal_End = True                       '異常終了
    Unload Me
    


End Sub


Private Sub Command2_Click()


Dim sts     As Integer
Dim yn      As Integer

    If Not IsNumeric(Text2(0).Text) Or Not IsNumeric(Text2(1).Text) Then
        Label3.Caption = "「終了時刻」入力エラー"
        Text2(0).SetFocus
        Exit Sub
    End If

    Text2(0).Text = Format(Val(Text2(0).Text), "00")
    If Val(Text2(0).Text) < 0 Or Val(Text2(0).Text) > 23 Then
        Label3.Caption = "「終了時刻」入力エラー"
        Text2(0).SetFocus
        Exit Sub
    End If
    Text2(1).Text = Format(Val(Text2(1).Text), "00")
    If Val(Text2(1).Text) < 0 Or Val(Text2(1).Text) > 59 Then
        Label3.Caption = "「終了時刻」入力エラー"
        Text2(1).SetFocus
        Exit Sub
    End If
    
    sts = WriteIni(App.EXEName, "ENDTIME", "SYS", Text2(0).Text & Text2(1).Text)
    If sts Then
        Label3.Caption = "「終了時刻」書き込みエラー SYS.INI"
    End If
End Sub

'[2014/02/10 - M.MATSUYAMA 変更(Ver2.0.0)] ソケット通信用追加
'Private Sub CtrsWsk1_OnReceive(ByVal ID_NO As Integer, ByVal RecvText As String, ByVal Resp_Mode As Boolean)
Private Sub OnDataReceive(ByVal strPID As String, ByVal strTID As String, ByVal ID_NO As Integer, ByVal intIndex As Integer, ByVal RecvText As String)
'-------------------------------------------------------
'
'   『レコード受信時処理』
'
'-------------------------------------------------------

Dim nErrCode    As Integer
Dim strErrMsg   As String
Dim intLine     As Integer
Dim i           As Integer
Dim j           As Integer
Dim Chk_Time    As String * 8
Dim Sendbuf     As String

Dim Errbuf      As String

Dim sts         As Integer

Dim Start_Flg   As Integer


Dim wkTEXT      As String

Dim Log_Out_txt As String       '2014.03.19
Dim k           As Integer      '2014.03.19

Dim l           As Integer      '2017.10.30


Dim wkHex       As String       '2017.09.07

    Text1(0).Text = Format(ID_NO, "000") & ", Recv=" & RecvText
    
    
    
    If F110010_LOG <> "" Then                                                                                           '2014.03.19
        Call LOG_OUT(F110010_LOG, "HT-->PC     " & Format(ID_NO, "000") & " " & Left(RecvText, Len(RecvText) - 2))      '2014.03.19
    End If                                                                                                              '2014.03.19
            
    RecvText = Left(RecvText, Len(RecvText) - 2)
    
                                    'ＩＤ№で受信済みテキスト検索
    ING_No = -1
    
    Start_Flg = False
    
    For i = 0 To UBound(ID_KANRI_TBL)
        If ID_NO = ID_KANRI_TBL(i).ID Then
            ING_No = i
            Chk_Time = ID_KANRI_TBL(i).Time
            Exit For
        End If
    Next i
    
    
    
    
    
    If i > UBound(ID_KANRI_TBL) Then
                                                'ＩＤ№新規登録
        For i = 0 To UBound(ID_KANRI_TBL)
            If ID_KANRI_TBL(i).ID = 0 Then
                
                Start_Flg = True
                
                ID_KANRI_TBL(i).ID = ID_NO      'ID_No  保存
                
'                ID_KANRI_TBL(i).MENU_GRP = ""
                ID_KANRI_TBL(i).MENU_LV1 = ""
                ID_KANRI_TBL(i).MENU_LV2 = ""
''                ID_KANRI_TBL(i).MENU_LV3 = ""
                
                If UBound(JGYOBU_T) = 0 Then    '１事業部固定
                Else
                    ID_KANRI_TBL(i).JGYOBU = ""
                End If
                
                If UBound(NAIGAI) = 0 Then   '国内外固定
                Else
                    ID_KANRI_TBL(i).NAIGAI = ""
                End If
                
                ING_No = i
                Chk_Time = ""
                Exit For
            End If
        
        Next i
    End If
    
    
'Call Log_Out(LOG_F, Format(ID_NO, "000") & ",Yoin= " & ID_KANRI_TBL(i).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(i).Sagyo_Code.YOIN_CODE)
    
    
    If ING_No = -1 Then
        MsgBox "ＩＮＩファイルの子機数の設定を変更して下さい。", vbCritical
        Normal_End = True
        Unload Me
    End If
    
                                            '前回受信値を再受信した？
    If Left(Right(RecvText, 9), 8) = ID_KANRI_TBL(i).Time And _
         Right(RecvText, 1) = "1" Then
            Call Send_Err_Proc(Sendbuf)
    
            Call LOG_OUT(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf & "[再送信]")
If F110010_LOG <> "" Then                                                                                                                                       '2014.03.19
    Call LOG_OUT(F110010_LOG, "HT-->PC     " & Format(ID_NO, "000") & " " & Left(RecvText, Len(RecvText) - 2) & "[再受信]前:" & ID_KANRI_TBL(ING_No).Time)      '2014.03.19
End If                                                                                                                                                          '2014.03.19
    Else
                                            
        '>>>>>>>>>>>>>>>>>>>>>> 2013.01.04
        
        If Val(Left(Right(RecvText, 9), 8)) <> 0 Then
            If Left(Right(RecvText, 9), 8) < ID_KANRI_TBL(ING_No).Time Then
                                                    
                Call Err_Send_Proc("受信エラー", "再読込みして下さい", "", "", "")
                Sendbuf = Text_Create_Proc()
If F110010_LOG <> "" Then                                                                                                                                       '2014.03.19
    Call LOG_OUT(F110010_LOG, "HT-->PC     " & Format(ID_NO, "000") & " " & Left(RecvText, Len(RecvText) - 2) & "[受信エラー]前:" & ID_KANRI_TBL(ING_No).Time)  '2014.03.19
End If                                                                                                                                                          '2014.03.19
                GoTo SendResp_Proc
            End If
        End If
        '>>>>>>>>>>>>>>>>>>>>>> 2013.01.04
                                            '受信内容を保存
        
        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_LotNo_LABEL_PRINT Then   '2013.06.06
        
            ID_KANRI_TBL(ING_No).Recv_text(0) = f_MidB(RecvText, 1, 20)
            ID_KANRI_TBL(ING_No).Recv_text(1) = f_MidB(RecvText, 21, 20)
            ID_KANRI_TBL(ING_No).Recv_text(2) = f_MidB(RecvText, 41, 20)
            ID_KANRI_TBL(ING_No).Recv_text(3) = f_MidB(RecvText, 61, 20)
            ID_KANRI_TBL(ING_No).Recv_text(4) = f_MidB(RecvText, 81, 20)
        
        Else                                                                                                                    '2013.06.06
        
            ID_KANRI_TBL(ING_No).Recv_text(0) = Left(RecvText, 20)       '受信内容１行目
            
            
            
'>>>>>>>>>>>>2017.09.07
            
'            For j = 0 To UBound(ID_KANRI_TBL(ING_No).Send_Text.Box_Type)
'
'                If Mid(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18, 1, 2) = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 1, 2) Then
'
'                    ID_KANRI_TBL(ING_No).Recv_text(0) = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18 = ""
'                    Exit For
'                End If
'            Next j
            
'>>>>>>>>>>>>2017.09.07
            
            ID_KANRI_TBL(ING_No).Recv_text(1) = Mid(RecvText, 21, 20)    '受信内容２行目
            ID_KANRI_TBL(ING_No).Recv_text(2) = Mid(RecvText, 41, 20)    '受信内容３行目
            ID_KANRI_TBL(ING_No).Recv_text(3) = Mid(RecvText, 61, 20)    '受信内容４行目
            ID_KANRI_TBL(ING_No).Recv_text(4) = Mid(RecvText, 81, 20)    '受信内容４行目
        
        
        
        
        
        
        End If                                                                                                                  '2013.06.06
        
        ID_KANRI_TBL(ING_No).Time = Left(Right(RecvText, 9), 8)      '送信時刻
        

If F110010_LOG <> "" Then                                                                                           '2014.03.19
    Call LOG_OUT(F110010_LOG, "PC(処理)    " & Format(ID_NO, "000") & " 要因=" & ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE & " Step=" & ID_KANRI_TBL(ING_No).Step & " Time=" & ID_KANRI_TBL(ING_No).Time)   '2014.03.19
End If                                                                                                              '2014.03.19

        If Start_Flg Then
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> Start_Para Then
                Call Err_Send_Proc("再起動してください。", "", "", "", "")
                Sendbuf = Text_Create_Proc()
            End If
        End If
                                            
        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SHOUHINKA Then
            
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = Fin_Para Then
            
                If ID_KANRI_TBL(ING_No).Step >= Step_Sagyo2_REQ Then
                
                    ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
                End If
            End If
        End If
                                      
                                      
                                      
                                      
        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_COMPO_CHECK Then
            
            
            If ID_KANRI_TBL(ING_No).Step >= Step_Sagyo2_REQ Then
                
                If Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(ID_KANRI_TBL(ING_No).Input_Line + 2)), 2) = LCD_Hinban Then
                            
                
                
                
                    ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
                End If
            End If
        End If
                                      
                                      
                                      
        '2013.01.18
        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_SEK_PACKING_ALL Then
            If ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ Then
            
                ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
                                  
            End If
        End If
        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_DEN_ALL) Or _
            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_LOGISTIC_ALL) Or _
            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_E_BAG_ALL) Then
            If ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ Then
            
                ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
                                  
            End If
        End If
        '2013.01.18
                                      
                                      
        '2013.04.04
'        If ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ Then
'
'            ID_KANRI_TBL(ING_No).Recv_text(0) = "ENT"
'
'        End If
        '2013.04.04
                                      
                                      
                                      
'>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.21
        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = Wel_T_back Then
            If ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ Then
                            
                If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> Can_Para Then
                    ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
                End If
            End If
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>   2015.01.21
                                      
                                      
'>>>>>>>>>>>>>>>>>>>>>>>>   2016.10.14
        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE) = ACT_KENPIN_Drct Then
            If ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ Then
                ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
            End If
                 
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE ' 2020/03/10 荷積明細処理の為、追加
               Case "Q4"
                 If ID_KANRI_TBL(ING_No).Step = 25 Then
                    ID_KANRI_TBL(ING_No).Step = 26
                 ElseIf ID_KANRI_TBL(ING_No).Step = 27 Then '2020/03/10 予測
                    ID_KANRI_TBL(ING_No).Step = Ent_Para
                 End If
            End Select
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>   2016.10.14
        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = TANTO_REF Then
            ID_KANRI_TBL(ING_No).Recv_text(0) = Can_Para
            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
            ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>   2017.09.27


'>>>>>>>>>>>>>>>>>>>>>>>>   2018.10.03
        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = Wel_MODULE_IN Then
            
            If ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ Then
                ID_KANRI_TBL(ING_No).Recv_text(0) = Ent_Para
            End If
        End If
'>>>>>>>>>>>>>>>>>>>>>>>>   2018.10.03

                                      
                                      
                                      
                                      
                                      '[START][CANCEL][FINISHI]受信は初期化する
        Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
            Case Start_Para     '開始(子機電源ON)
                
                
                            '出荷予定／在庫の予約解除
                sts = Data_Clear_Proc(0, Sendbuf)
                Select Case sts
                    Case SYS_ERR
                        Normal_End = True
                End Select
                
                
                ID_KANRI_TBL(ING_No).Step = Step_Start
        
                ID_KANRI_TBL(ING_No).CTR_TYPE = Trim(ID_KANRI_TBL(ING_No).Recv_text(1))
                        
        
                Call Start_Proc(Sendbuf)
            
            
            Case Ent_Para       'ENT
                If Not Start_Flg Then
                                                                    '2016.10.14 ACT_KENPIN_Drct追加
                    If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_Drct Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_DEN Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SPECIAL_PROCESS Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_BINNO Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_GAI Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SHOUHINKA Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_LotNo Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_MODULE Then '2013.06.06-->2014.06.24
                        
                        
                        '検品時の確認
                        
                        
                        
    '                    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    '                    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    '                    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    '                    Select Case sts
    '                        Case BtNoErr
    '                        '   -------------------------------- エラーメッセージ作成
    '                        Case Else
    '                       '重要な要因なので未登録はシステム停止とする
    '                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
    '                        Sendbuf = Text_Create_Proc()
    '                        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
    '                        Normal_End = True
    '                    End Select
    '
                        
                        
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                        If Sagyo_Main_Proc(Sendbuf) Then
                            Normal_End = True
    '                        Unload Me
                        End If
                    Else
                        
                        '--------------------------------------------------- 大阪  部材対応　2012.03.06
                        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_HIN_FURIKA_S Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_IDO_IN_OSAKA Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_IDO_OUT_OSAKA2 Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_IDO_OUT_OSAKA3 Then     '206.05.11
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                            
                            If Sagyo_Main_Proc(Sendbuf) Then
                                Normal_End = True
                            End If
                        
                        Else
                        '--------------------------------------------------- 大阪  部材対応　2012.03.06
                            '参照画面の確認時のみ
                            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                            Select Case sts
                                Case BtNoErr
                                 '   -------------------------------- エラーメッセージ作成
                                Case Else
                               '重要な要因なので未登録はシステム停止とする
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                                Normal_End = True
                            End Select
                            
                             
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Normal_End = True
                            End If
                            Sendbuf = Text_Create_Proc()
                        End If
                    End If
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            Case Can_Para       'CANCEL
                If Not Start_Flg Then
                
                    If ID_KANRI_TBL(ING_No).Last_Send = 1 Then
                                
If 0 = 100 Then
    GoTo print_end
End If
                        
                        '検品時はデータの開放を行う　2004.06.14 ↓
                        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SPECIAL_PROCESS Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_BINNO Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_GAI Then '2009.08.08


                                 
                            
                            
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    If Sagyo_Send_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                    Sendbuf = Text_Create_Proc()
                                
                                Case SYS_ERR
                                    Normal_End = True
                            End Select
                        
                        Else
                        
'''                             大阪PC向けは、特定のｽﾃｯﾌﾟのみ解除する
                            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_DEN Then '2006.12.07
                                If (ID_KANRI_TBL(ING_No).Step <> Step_Sagyo2_RES And _
                                    ID_KANRI_TBL(ING_No).Step <> Step_Sagyo3_RES And _
                                    ID_KANRI_TBL(ING_No).Step <> Step_Sagyo5_RES) Then
                                
                                    sts = Data_Clear_Proc(0, Sendbuf)
                                    Select Case sts
                                        Case SYS_CANCEL
                                            If Sagyo_Send_Proc() Then
                                                Sendbuf = Text_Create_Proc()
                                                Normal_End = True
                                            End If
                                            Sendbuf = Text_Create_Proc()
                                        
                                        Case SYS_ERR
                                            Normal_End = True
                                    End Select
                                
                                End If
                            End If
                        End If
                                
                        '検品時はデータの開放を行う　2004.06.14 ↑
                                
                        
                        
                        'ﾃﾞｰﾀ解放処理を追加 2011.04.07
                        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SYUKA_HYO Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_DENPYO_ID Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_DENPYO_ID2 Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SYUKA_HYO_OSAKA) Then       'ACT_DENPYO_ID2を追加 2015.02.21
                            If (ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ) Then


                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        If Sagyo_Send_Proc() Then
                                            Sendbuf = Text_Create_Proc()
                                            Normal_End = True
                                        End If
                                        Sendbuf = Text_Create_Proc()

                                    Case SYS_ERR
                                        Normal_End = True
                                End Select

                            End If
                        End If
                        'ﾃﾞｰﾀ解放処理を追加 2011.04.07
                                
                                
                                '前回がエラー送信
                        Call Re_Send_Proc(Sendbuf)
                
                    Else
                        
                        '2015.01.21
                        If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = Wel_T_back And _
                            ID_KANRI_TBL(ING_No).Step >= Step_Sagyo3_REQ Then
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                            
                            Call T_back_Step_1_Proc
                            Sendbuf = Text_Create_Proc()
                        Else
                        '2015.01.21
                        
                            If ID_KANRI_TBL(ING_No).Step = Step_Check_REQ Then      '2013.07.25
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ         '2013.07.25
                            End If                                                  '2013.07.25
                            
                            
                            '>>>>>>>>>>>>>>>>>>>>>  2013.04.05
                            
                            
                            If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_DEN_ALL) Or _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_LOGISTIC_ALL) Or _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = Wel_Inspe_E_BAG_ALL) Then
                                If ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ Then
                                    If Print_Cancel_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                End If
                            End If
                            '>>>>>>>>>>>>>>>>>>>>>  2013.04.05
                                    
                                    '出荷予定／在庫の予約解除
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    If Sagyo_Send_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                    Sendbuf = Text_Create_Proc()
                                
                                Case SYS_ERR
                                    Normal_End = True
                            End Select
                    
                    
                    
                            Call Cancel_Proc(Sendbuf)
                        
                        End If          '2015.01.21
                
                    End If
            
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            
            Case Fin_Para       'FINISH
                    
                If Not Start_Flg Then
                
                
                
                    If ID_KANRI_TBL(ING_No).LABEL_ON Then
                        ID_KANRI_TBL(ING_No).LABEL_ON = False
                
                
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                    
                        If Sagyo_Main_Proc(Sendbuf) Then
                            Normal_End = True
                        End If
                
                    Else
                
                
                        '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                If Sagyo_Send_Proc() Then
                                    Normal_End = True
                                End If
                                Sendbuf = Text_Create_Proc()
                        
                            Case SYS_ERR
                                Normal_End = True
                        End Select
                    
                    
    '                    If Step_MENU1_REQ < ID_KANRI_TBL(ING_No).Step Then
                        If Step_TANTO_REQ <> ID_KANRI_TBL(ING_No).Step Then      '2005.01.07 if ～　else ～　end if
                        
                            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
    '2006.01.30                        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                    
                            ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
    '2006.01.03                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                    
                    
                    
                            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    
                            If Menu_Send_Proc(Sendbuf) Then
                                Normal_End = True
        '                Unload Me
                            End If
                    
                        Else                                                    '2005.01.07
    '                        ID_KANRI_TBL(ING_No).Step = Step_Start
                                                                                '2005.01.07
                            Call Start_Proc(Sendbuf)                            '2005.01.07
                                                                                '2005.01.07
                        End If                                                  '2005.01.07
                    End If
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            
            
            '2008.08.08 MENU FIN
            Case MENU_FIN
            
                            
                If Not Start_Flg Then
                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                ID_KANRI_TBL(ING_No).NAIGAI, _
                                                ID_KANRI_TBL(ING_No).MENU_LV1, _
                                                "EN", , , , , , , , , FILE_RETRY) Then
                                    
                        Normal_End = True
                    
                    End If
                            
                    Text1(1).Text = Format(ID_KANRI_TBL(ING_No).ID, "000") & "=" & "EN"
                            
                            
'                    ST_LOG_OUT_F = True
                
                    If ID_KANRI_TBL(ING_No).Last_Send = 1 Then
                                
                                
                        '検品時はデータの開放を行う　2004.06.14 ↓
                        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_NEW_KENPIN Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_NEW_KENPIN_MTS Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_SPECIAL_PROCESS Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_BINNO Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_GAI Then '2009.08.08


                                 
                            
                            
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    If Sagyo_Send_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                    Sendbuf = Text_Create_Proc()
                                
                                Case SYS_ERR
                                    Normal_End = True
                            End Select
                        
                        Else
                        
'''                             大阪PC向けは、特定のｽﾃｯﾌﾟのみ解除する
                            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_DEN Then '2006.12.07
                                If (ID_KANRI_TBL(ING_No).Step <> Step_Sagyo2_RES And _
                                    ID_KANRI_TBL(ING_No).Step <> Step_Sagyo3_RES And _
                                    ID_KANRI_TBL(ING_No).Step <> Step_Sagyo5_RES) Then
                                
                                    sts = Data_Clear_Proc(0, Sendbuf)
                                    Select Case sts
                                        Case SYS_CANCEL
                                            If Sagyo_Send_Proc() Then
                                                Sendbuf = Text_Create_Proc()
                                                Normal_End = True
                                            End If
                                            Sendbuf = Text_Create_Proc()
                                        
                                        Case SYS_ERR
                                            Normal_End = True
                                    End Select
                                
                                End If
                            End If
                        End If
                                
                        '検品時はデータの開放を行う　2004.06.14 ↑
                                
                                '前回がエラー送信
                        Call Re_Send_Proc(Sendbuf)
                
                    Else
                                '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                If Sagyo_Send_Proc() Then
                                    Sendbuf = Text_Create_Proc()
                                    Normal_End = True
                                End If
                                Sendbuf = Text_Create_Proc()
                            
                            Case SYS_ERR
                                Normal_End = True
                        End Select
                
                
                
                        Call Cancel_Proc(Sendbuf, 1, "EN")          '2008.09.01 ﾊﾟﾗﾒｰﾀ追加
                
                    End If
            
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            
            
            
            '2008.08.08 MENU FIN
            
            
            
            
            
            Case P_END_Para     '印刷完了   2010.01.21
            
print_end:
            
                ID_KANRI_TBL(ING_No).LABEL_ON = False
        
        
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
            
                If Sagyo_Main_Proc(Sendbuf) Then
                    Normal_End = True
                End If
            
            '▼[2014/02/20 - M.MATSUYAMA 変更(Ver2.0.0)] fileloadを廃止しソケット通信にてファイル情報を送信する
            Case P_FILELOAD
                Dim strFileName As String   '送信ファイル名
                Dim strFileData As String   '全ファイルデータ
                Dim strMsg As String        'エラーメッセージ
                Dim lngFileSize As Long     '全ファイルサイズ
                
                'ファイル名を取得
                strFileName = Trim(ID_KANRI_TBL(ING_No).Recv_text(1))
                'ファイル名がセットされていたら
                If Len(strFileName) > 0 Then
                    '全ファイルデータ取得
                    strFileData = GetFileData(GetFullPath(CtrsWsk1.SendFolder, strFileName))
                    '全ファイルサイズ取得
                    lngFileSize = LenB(StrConv(strFileData, vbFromUnicode))
                    'ファイルサイズが0以上の場合
                    If lngFileSize > 0 Then
                        '最大送信データサイズ以内の場合、電文に付加する
                        If lngFileSize <= MAX_FSENDDATASIZE Then
                            Sendbuf = RESP_OK & Format(lngFileSize, "0000") & strFileData & vbCrLf
                        Else
                            strMsg = AlignText("表示件数が20件を超えました(棚番または移動歴)", 60, vbLeftJustify)  '2020/04/03 エラーメッセージ変更
                            Sendbuf = RESP_NG & "<ファイル受信エラー>" & strMsg & vbCrLf
                        End If
                    Else
                        strMsg = AlignText("ファイルサイズが０バイトです", 60, vbLeftJustify)
                        Sendbuf = RESP_NG & "<ファイル受信エラー>" & strMsg & vbCrLf
                    End If
                Else
                    strMsg = AlignText("ファイル名がセットされていません", 60, vbLeftJustify)
                    Sendbuf = RESP_NG & "<ファイル受信エラー>" & strMsg & vbCrLf
                End If
            '▲[2014/02/20 - M.MATSUYAMA 変更(Ver2.0.0)] fileloadを廃止しソケット通信にてファイル情報を送信する
            
            Case Else
                If Not Start_Flg Then
                                            '進捗チェック
                    Select Case ID_KANRI_TBL(ING_No).Step
            
            
                        Case Step_TANTO_REQ         '担当者要求に対するレス
                        
                            
                            
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_RES
        
                            If Normal_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
        
                        Case Step_JGYOBU_REQ        '事業部要求に対するレス
                
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
    '                        Case BEF_Page       '前頁
    '                        Case NEXT_Page      '次頁
                                Case Else            '事業部受信
                
                                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_RES
              
                                    ID_KANRI_TBL(ING_No).JGYOBU = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                                    
                        Case Step_NAIGAI_REQ
                    
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case Else           'メニューパラメータ受信
                
                                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_RES
                                        
                                    ID_KANRI_TBL(ING_No).NAIGAI = Trim(ID_KANRI_TBL(i).Recv_text(0))
                                
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
    
'2006.01.30                        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
                        Case Step_MENU1_REQ, Step_MENU2_REQ
                
                             Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case BEF_Page       '前頁
                            
                                    
                                    
                                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.02.22
                                    If ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU And ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head Then
                                        Call LOG_OUT(LOG_F, "ﾒﾆｭｰ指示異常受信 " & RecvText)

                                        Call Re_Send_Proc(Sendbuf)
                                        GoTo SendResp_Proc
                                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.02.22

                                    
                                    
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 - 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 - 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 - 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                                Case NEXT_Page      '次頁
                                    
                                    
                                    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.02.22
                                    If ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU And ID_KANRI_TBL(ING_No).Send_Text.End_Menu = MENU_END Then
                                        Call LOG_OUT(LOG_F, "ﾒﾆｭｰ指示異常受信 " & RecvText)

                                        Call Re_Send_Proc(Sendbuf)
                                        GoTo SendResp_Proc
                                    End If
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>   2017.02.22
                                    
                                    
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 + 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 + 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 + 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                Case Else           'メニューパラメータ受信
                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                                        Case Step_MENU2_REQ
'                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                                    End Select
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV1 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                        
                                        
                                            '2008.08.08
                                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                                        Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                                                        ID_KANRI_TBL(ING_No).MENU_LV1, _
                                                                        "ST", , , , , , , , , FILE_RETRY) Then
                            
                                                Normal_End = True
                                            End If
                                            
                                                                                
                                            Text1(1).Text = Format(ID_KANRI_TBL(ING_No).ID, "000") & "=" & "ST"
                                            
                                            
                                            '2008.08.08
                                        
                                        
                                        
                                        Case Step_MENU2_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV2 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                            'ID_KANRI_TBL(ING_No).MTS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8)                   '2017.09.07
                                            '>>>>>>>>>>>>   2017.10.30
                                            wkHex = ""
                                            For l = 0 To M_Gyo - 1
                                                If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = Trim(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(l).MENU) Then
                                                   If Trim(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(l).MTS_FLG) = "" Then
                                                    
                                                        wkHex = f16sinTo10sin(Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8))                             '2017.09.07
                                                        Exit For
                                                    End If
                                                End If
                                            Next l
                                            '>>>>>>>>>>>>   2017.10.30
                                            
                                            '2017.10.30wkHex = f16sinTo10sin(Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8))                             '2017.09.07
                                            
                                            If Trim(wkHex) = "" Then                                                                        '2017.09.07
                                                ID_KANRI_TBL(ING_No).MTS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8)                '2017.09.07
                                            Else                                                                                            '2017.09.07
                                                ID_KANRI_TBL(ING_No).MTS_CODE = wkHex                                                       '2017.09.07
                                            End If
                                            
                                            If IsNumeric(ID_KANRI_TBL(ING_No).MTS_CODE) Then                                                '2017.09.07
                                                ID_KANRI_TBL(ING_No).MTS_CODE = Format(Val(ID_KANRI_TBL(ING_No).MTS_CODE), "00000000")      '2017.09.07
                                            End If                                                                                          '2017.09.07
                                            ID_KANRI_TBL(ING_No).SS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 11, 8)

'2006.01.30                                        Case Step_MENU3_RES
'2006.01.30                                            ID_KANRI_TBL(ING_No).MENU_LV3 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                    
                                    
                                    
                                            '2008.08.08
'                                            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
'                                                                        Format(ID_KANRI_TBL(ING_No).ID, "000"), _
'                                                                        ID_KANRI_TBL(ING_No).JGYOBU, _
'                                                                        ID_KANRI_TBL(ING_No).NAIGAI, _
'                                                                        ID_KANRI_TBL(ING_No).MENU_LV1, _
'                                                                        "ST", , , , , , , , , FILE_RETRY) Then
'
'                                                Normal_End = True
'                                            End If
                                            '2008.08.08
                                    
                                    
                                    
                                    
                                    
                                    
                                    End Select
                
                                
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                        
                        Case Step_Sagyo1_REQ, Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ, Step_Sagyo5_REQ, Step_Sagyo6_REQ
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                            If Sagyo_Main_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
                                                
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.07.25
                        Case Step_Check_REQ
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                            If Sagyo_Main_Proc(Sendbuf) Then
                                Normal_End = True
                            End If
                        '>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.07.25
                        
                        
                        Case Else
                            Call Err_Send_Proc("受信エラー", "再読込みして下さい", "", "", "")                      '2014.03.19
                            Sendbuf = Text_Create_Proc()                                                            '2014.03.19
                            GoTo SendResp_Proc                                                                      '2014.03.19
                    End Select
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
    

    
        
        End Select
    
    End If
    
    
SendResp_Proc:              '2013.01.04

    '2014.03.19 ↓
    Log_Out_txt = Send_Text.sts & Send_Text.Display_Flg & Send_Text.End_Menu & Send_Text.Menu_Suu & Send_Text.FileName & Send_Text.buzzer

    For k = 0 To M_Gyo - 1
                                                                                    'BOX属性
            
            Log_Out_txt = Log_Out_txt & Send_Text.Box_Type(k).Box_Type
                                                                                    '表示内容
            Log_Out_txt = Log_Out_txt & StrConv(Send_Text.Box_Type(k).LCD, vbUnicode)
                                                                                    '初期内容
            Log_Out_txt = Log_Out_txt & Send_Text.Box_Type(k).INIT
                                                                                    '開始カーソル位置
            Log_Out_txt = Log_Out_txt & Send_Text.Box_Type(k).Start_Pos
                                                                                    '入力桁数（最大）
            Log_Out_txt = Log_Out_txt & Send_Text.Box_Type(k).Max_Size
                                                                                    'メニュー内容
            Log_Out_txt = Log_Out_txt & Send_Text.Box_Type(k).MENU
        
    Next k

    If F110010_LOG <> "" Then
        Call LOG_OUT(F110010_LOG, "PC-->HT BEF " & Format(ID_NO, "000") & " " & Log_Out_txt)
    End If
    '2014.03.19 ↑



    '[2014/02/10 - M.MATSUYAMA 変更(Ver2.0.0)] ソケット通信用追加
    'If Resp_Mode Then
    If strPID = "R" Then
        On Error GoTo ShowError

        '▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
        'CtrsWsk1.SendResp Sendbuf
        If intIndex > 0 Then
            Dim strRespHeader As String 'ヘッダ情報
            Dim strName As String
            
            strName = IIf(Len(tcpHost(intIndex).RemoteHost) = 0, _
                            tcpHost(intIndex).RemoteHostIP, _
                            tcpHost(intIndex).RemoteHost & " (" & tcpHost(intIndex).RemoteHostIP & ")")
            
            'レスポンスデータセット
            strRespHeader = strTID & "A"
            gbl_RespBuf(intIndex) = strRespHeader + Sendbuf
            Call WriteLogMsg("[" & ConvBinaryMsg(gbl_RespBuf(intIndex)) & "]", FNC_SOCKSEND, ID_NO, strName, icoUpload)
            'データ送信
            Call tcpHost(intIndex).SendData(gbl_RespBuf(intIndex))
        Else
            Call WriteLogMsg("[" & ConvBinaryMsg(Sendbuf) & "]", FNC_SENDDATA, ID_NO, , icoUpload)
            CtrsWsk1.SendResp Sendbuf
        End If
        '▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加

'        Text1(1).Text = Format(ID_NO, "000") & ", Send=" & SendBuf
        
'        Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)

        On Error GoTo 0
    
        ID_KANRI_TBL(ING_No).Last_Send_Text.sts = Send_Text.sts                     'ステータス
        ID_KANRI_TBL(ING_No).Last_Send_Text.Display_Flg = Send_Text.Display_Flg     '表示画面フラグ
        ID_KANRI_TBL(ING_No).Last_Send_Text.End_Menu = Send_Text.End_Menu           '最終メニューフラグ
        ID_KANRI_TBL(ING_No).Last_Send_Text.Menu_Suu = Send_Text.Menu_Suu           'メニュー個数
        ID_KANRI_TBL(ING_No).Last_Send_Text.FileName = Send_Text.FileName           'ファイル名
        ID_KANRI_TBL(ING_No).Last_Send_Text.buzzer = Send_Text.buzzer               'ブザー指定
        
        For j = 0 To M_Gyo - 1
                                                                                    'BOX属性
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Box_Type = Send_Text.Box_Type(j).Box_Type
                                                                                    '表示内容
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).LCD, StrConv(Send_Text.Box_Type(j).LCD, vbUnicode))
                                                                                    '初期内容
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).INIT = Send_Text.Box_Type(j).INIT
                                                                                    '開始カーソル位置
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Start_Pos = Send_Text.Box_Type(j).Start_Pos
                                                                                    '入力桁数（最大）
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Max_Size = Send_Text.Box_Type(j).Max_Size
                                                                                    'メニュー内容
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).MENU = Send_Text.Box_Type(j).MENU
            
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).MENU18 = Send_Text.Box_Type(j).MENU18   '2017.09.07
                    
        
        Next j
    
    
        If Normal_End Then
            
'            MsgBox "システム異常が発生しました！！処理をしてください。"
 
            
'            Unload Me
        End If
    End If

    If F110010_LOG <> "" Then                                                                   '2014.03.19

Call LOG_OUT(F110010_LOG, "PC-->HT AFT " & Format(ID_NO, "000") & " " & Log_Out_txt)     '2014.03.19
    
    
'>>>>>>>>>>>>>>>>>>>>>> 2016.01.21
        If Trim(ID_KANRI_TBL(ING_No).Last_Send_Text.FileName) <> "" Then                                                                                                                                       '2014.03.19
            strFileName = ID_KANRI_TBL(ING_No).Last_Send_Text.FileName
            strFileData = GetFileData(GetFullPath(CtrsWsk1.SendFolder, strFileName))
            lngFileSize = LenB(StrConv(strFileData, vbFromUnicode))
            If lngFileSize > 0 Then
                Call LOG_OUT(F110010_LOG, "PC-->HT BEF " & Format(ID_NO, "000") & " " & strFileName & Chr(&HD) & Chr(&HA) & strFileData)
            Else
                Call LOG_OUT(F110010_LOG, "PC-->HT BEF " & Format(ID_NO, "000") & " " & strFileName & " " & "Nodata")
            End If
        End If                                                                                                                                                          '2014.03.19
'>>>>>>>>>>>>>>>>>>>>>> 2016.01.21
    
    
    
    
    End If                                                                                      '2014.03.19


    Exit Sub

ShowError:
    nErrCode = Err.Number
    strErrMsg = Err.Description         'エラーメッセージ
    
    intLine = CtrsWsk1.ErrLineNo        '接続番号を取得します。
    If intLine > 0 Then
        strErrMsg = strErrMsg & Chr(&HD) & Chr(&HA) & "接続番号 = " & intLine
    End If

    Text1(2).Text = strErrMsg           'ステータス行にエラーを表示します。
    
'    Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)


End Sub

Private Sub Form_Activate()
    
'    Command1(0).Value = True            '2012.11.06        2013.05.22
'    Timer1.Enabled = True               '2012.11.06        2013.05.22

End Sub

Private Sub Form_Load()
    
Dim c           As String * 128
Dim Out_Data    As String

Dim Box_Type    As String * 1
Dim LCD         As String * 20
Dim Keta        As String * 2

Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim sts         As Integer
    
Dim wkSHIMUKE   As Variant
    
Dim wkUNSOU_KAISHA  As Variant      '運送会社読み替え　2006.12.22
    
Dim SHIMEBI     As String * 2       '2012.03.06
Dim wkYY        As Integer          '2012.03.06
Dim wkMM        As Integer          '2012.03.06
    
Dim wkTH        As String * 4       '2012.11.06
    
    
Dim wkVariant   As Variant          '2014.07.01
    
    
    Normal_End = False
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
'---------------------------------------------- 'ログファイル名取り込み
    
    
    
    
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)

'---------------------------------------------- 'データ受信ポート番号取り込み
    If GetIni(App.EXEName, "LocalPort", "SYS", c) Then
        Beep
        MsgBox "データ受信ポート番号の獲得に失敗しました。処理を中止します。"
        End
    End If
    LocalPort = CLng(RTrim(c))


    lblINI(1).Caption = "    LocalPort:" & RTrim(c)     '2014.07.01
'---------------------------------------------- 'データ送信ポート番号取り込み
    If GetIni(App.EXEName, "RemotePort", "SYS", c) Then
        Beep
        MsgBox "データ送信ポート番号の獲得に失敗しました。処理を中止します。"
        End
    End If
    RemotePort = CLng(RTrim(c))
    lblINI(2).Caption = "   RemotePort:" & RTrim(c)    '2014.07.01
'---------------------------------------------- '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
'---------------------------------------------- '国内外情報取り込み
    i = 0
    
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI_CODE" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI(i - 1)
        NAIGAI(i - 1).CODE = Trim(c)
        If GetIni(App.EXEName, "NAIGAI_NAME" & Format(i, "0"), "SYS", c) Then
            MsgBox "国内外の獲得に失敗しました。処理を中止します。"
            End
        End If
        NAIGAI(i - 1).NAME = Trim(c)
    
    Loop
    
    If i = 1 Then
        Beep
        MsgBox "国内外情報の獲得に失敗しました。処理を中止します。"
        End
    End If
'---------------------------------------------- 仕向け先獲得
    Erase SHIMUKE_TBL

    i = 0
    Do
        
        i = i + 1
    
        If GetIni(App.EXEName, "SHIMUKE" & Format(i, "00"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve SHIMUKE_TBL(i - 1)
        wkSHIMUKE = Split(c, ",", -1)
        If UBound(wkSHIMUKE) < 2 Then
            MsgBox "仕向け先情報の獲得に失敗しました。処理を中止します。"
        End If
    
        SHIMUKE_TBL(i - 1).JGYOBU = wkSHIMUKE(0)
        SHIMUKE_TBL(i - 1).NAIGAI = wkSHIMUKE(1)
        SHIMUKE_TBL(i - 1).SHIMUKE_CODE = wkSHIMUKE(2)
    Loop
'---------------------------------------------- 前借り入荷情報獲得
    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_MAEGARI = Trim(c)
'---------------------------------------------- '国内外振替情報獲得
    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
    Else
        YOIN_FURIKAE = RTrim(c)
        '国内外振替設定時、以下の項目必須
        If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
            Beep
            MsgBox "国内外振替情報[YOIN_FURIKAE_OUT]の獲得に失敗しました。処理を中止します。"
            End
        End If
    
        YOIN_FURIKAE_OUT = RTrim(c)
    
        If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
            Beep
            MsgBox "国内外振替情報[YOIN_FURIKAE_IN]の獲得に失敗しました。処理を中止します。"
            End
        End If
    
        YOIN_FURIKAE_IN = RTrim(c)
    
    End If
'---------------------------------------------- 棚照合情報獲得
    If GetIni("YOIN", "YOIN_WEL_TANASHOGO", "SYS", c) Then
        Call LOG_OUT(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANASHOGO] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_TANASHOGO = Trim(c)

'---------------------------------------------- 棚品照合情報獲得
    If GetIni("YOIN", "YOIN_WEL_TANAHINSHOGO", "SYS", c) Then
        YOIN_TANAHINSHOGO = Wel_TANA_HIN_SHOGO
    Else
        YOIN_TANAHINSHOGO = Trim(c)
    End If



'---------------------------------------------- 品番照合情報獲得    2011.02.03
    If GetIni("YOIN", "YOIN_WEL_HIN_SHOGO", "SYS", c) Then
        YOIN_HIN_SHOGO = Wel_HIN_SHOGO
    Else
        YOIN_HIN_SHOGO = Trim(c)
    End If


'---------------------------------------------- 入荷倉庫    2007.06.07
    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "入荷用仮想倉庫番号の獲得に失敗しました。処理を中止します。"
        End
    Else
        NYUKA_SOKO_NO = Trim(c)
    End If
'---------------------------------------------- 大東倉庫№    2007.06.07
    If GetIni(App.EXEName, "SOKO_NO", "SYS", c) Then
        DAITO_SOKO_NO = "S1"
    Else
        DAITO_SOKO_NO = Trim(c)
    End If
'---------------------------------------------- 資材振替要因    2007.06.28
    If GetIni(App.EXEName, "FURIKAE", "SYS", c) Then
'        Beep
'        MsgBox "入荷用仮想倉庫番号の獲得に失敗しました。処理を中止します。"
'        End
    
        Wel_FURIKAE = ""
    
    Else
        Wel_FURIKAE = Trim(c)
    End If


'---------------------------------------------- 品番振替＋要因  2011.06.01
    If GetIni(App.EXEName, "HINBAN_FURIKAE_PLUS", "SYS", c) Then
    
        Wel_HIN_FURIKAE_PLUS = ""
    
    Else
        Wel_HIN_FURIKAE_PLUS = Trim(c)
    End If




    If Trim(Wel_HIN_FURIKAE_PLUS) <> "" Then
    
    
    '---------------------------------------------- 品番振替－要因  2011.06.01
        If GetIni(App.EXEName, "HINBAN_FURIKAE_MAINA", "SYS", c) Then
            Beep
            MsgBox "品番振替　振替－要因の獲得に失敗しました[F110010][HINBAN_FURIKAE_MAINA=]。処理を中止します。"
            End
        Else
            Wel_HIN_FURIKAE_MAINA = Trim(c)
        End If
    '---------------------------------------------- 品番振替－の移動歴出力時の要因
        If GetIni(App.EXEName, "BEF_HINBAN", "SYS", c) Then
            Beep
            MsgBox "品番振替　振替－の移動歴出力要因の獲得に失敗しました[F110010][BEF_HINBAN=]。処理を中止します。"
            End
        Else
            YOIN_BEF_HINBAN = Trim(c)
        End If
    '---------------------------------------------- 品番振替＋の移動歴出力時の要因
        If GetIni(App.EXEName, "AFT_HINBAN", "SYS", c) Then
            Beep
            MsgBox "品番振替　振替＋移動歴出力要因の獲得に失敗しました[F110010][AFT_HINBAN=]。処理を中止します。"
            End
        Else
            YOIN_AFT_HINBAN = Trim(c)
        End If

    End If
'---------------------------------------------- 品番振替要因    2011.06.01





'---------------------------------------------- 資材消費要因    2007.10.02
    If GetIni(App.EXEName, "S_SHOUHI", "SYS", c) Then
        Wel_S_SHOUHI = ""
    
    Else
        Wel_S_SHOUHI = Trim(c)
    End If


'---------------------------------------------- 資材消費（新）要因    2015.02.21
    If GetIni(App.EXEName, "S_SHOUHI2", "SYS", c) Then
        Wel_S_SHOUHI2 = ""
    
    Else
        Wel_S_SHOUHI2 = Trim(c)
    End If



'---------------------------------------------- '子機台数取り込み
    If GetIni(App.EXEName, "KO_SU", "SYS", c) Then
        Beep
        MsgBox "子機台数の獲得に失敗しました。処理を中止します。"
        End
    End If
    ReDim ID_KANRI_TBL(0 To CInt(RTrim(c)) - 1)

    For i = 0 To UBound(ID_KANRI_TBL)
        ID_KANRI_TBL(i).ID = 0          'IDNoクリアー
        ID_KANRI_TBL(i).Step = 0        '進捗クリアー
    
    Next i


    lblINI(0).Caption = "ﾊﾝﾃﾞｨ最大台数:" & RTrim(c)
'---------------------------------------------- '送信用パラメータ取り込み
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
            WEL_Para_Tbl(i, j).Action = ""
        Next j
    Next i
    
    i = 0
    Do
        i = i + 1
        
        
        If GetIni("ACTION", "ACTION_CD" & Format(i, "00"), "SYS", c) Then
            Beep
            MsgBox "WELCAT送信用パラーメータ([ACTION] [ACTION_CD])の獲得に失敗しました。処理を中止します。"
            End
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
    
        j = 0
    
        Do
            j = j + 1
            If GetIni("ACTION", "ACTION_WEL_PARA" & Format(i, "00") & Format(j, "00"), "SYS", c) Then
                Beep
                MsgBox "WELCAT送信用パラーメータ([ACTION] [ACTION_WEL_PARA])の獲得に失敗しました。処理を中止します。"
                End
            End If
            If Trim(c) = "NON" Then
                Exit Do
            End If
        
            Call Data_Select(Trim(c), 1, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Action = Trim(Out_Data)
        
            Call Data_Select(Trim(c), 2, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).Box_Type = Trim(Out_Data)
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).LCD = ""
        
        
            k = 2
            Do
                
                k = k + 1
                
                If k > 14 Then
                    Exit Do
                End If
                
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Box_Type = Trim(Out_Data)
                
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                LCD = Trim(Out_Data)
            
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Keta = Trim(Out_Data)
            
                Select Case k
                    Case 5
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Keta = CInt(Keta)
                    Case 8
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Keta = CInt(Keta)
                    Case 11
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Keta = CInt(Keta)
                    Case 14
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Keta = CInt(Keta)
                
                End Select
            
            Loop
            
        
        Loop
    Loop
'---------------------------------------------- '対WELCAT　送受信ログファイル取り込み
    
    If GetIni(App.EXEName, "LOG_F", "SYS", c) Then
        CtrsWsk1.LogFile = ""
    Else
        CtrsWsk1.LogFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　データ送信用フォルダ取り込み
    If GetIni(App.EXEName, "SendFolder", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用フォルダ([SYS.INI][F110010] [SendFolder])の獲得に失敗しました。処理を中止します。"
        End
    Else
        CtrsWsk1.SendFolder = Trim(c)
        '2013.06.06 フォルダ有無
        On Error Resume Next
        ChDir CtrsWsk1.SendFolder
        
        Select Case Err.Number
            Case 0
'            Case 75
            Case Else
                MsgBox "WELCAT送信用フォルダ([SYS.INI][F110010] [SendFolder])を正しく設定して下さい(該当ﾌｫﾙﾀﾞなし)。処理を中止します。"
                End
        End Select
        '2013.06.06
    End If
'---------------------------------------------- '対WELCAT　棚番表示用データファイル名取り込み
    If GetIni(App.EXEName, "B1", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B1])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B1_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　出庫履歴用データファイル名取り込み
    If GetIni(App.EXEName, "B6", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B6])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B6_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　出荷推移用データファイル名取り込み
    If GetIni(App.EXEName, "B7", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B7])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B7_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　棚番表示(仮想優先)用データファイル名取り込み
    If GetIni(App.EXEName, "B9", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B9])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B9_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　構成表示用データファイル名取り込み
    If GetIni(App.EXEName, "BA", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [BA])の獲得に失敗しました。処理を中止します。"
        End
    Else
        BA_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　検品履歴データファイル名取り込み
    If GetIni(App.EXEName, "BB", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [BB])の獲得に失敗しました。処理を中止します。"
        End
    Else
        BB_SendFile = Trim(c)
    End If



'---------------------------------------------- '対WELCAT　当日出庫履歴データファイル名取り込み 2009.01.09
    If GetIni(App.EXEName, "BD", "SYS", c) Then
        BD_SendFile = ""
    Else
        BD_SendFile = Trim(c)
    End If


'---------------------------------------------- '対WELCAT　集約梱包残データファイル名取り込み 2010.02.15
    If GetIni(App.EXEName, "BF", "SYS", c) Then
        BF_SendFile = ""
    Else
        BF_SendFile = Trim(c)
    End If


'---------------------------------------------- '対WELCAT　空き棚検索データファイル名取り込み 2010.12.13
    If GetIni(App.EXEName, "BG", "SYS", c) Then
        BG_SendFile = ""
    Else
        BG_SendFile = Trim(c)
    End If


'---------------------------------------------- '対WELCAT　大阪検品データファイル名取り込み 2010.01.21
    If GetIni(App.EXEName, "F0", "SYS", c) Then
        F0_SendFile = ""
    Else
        F0_SendFile = Trim(c)
    End If



'---------------------------------------------- '対WELCAT　ﾗﾍﾞﾙ発行用データファイル名取り込み 2011.03.05
    If GetIni(App.EXEName, "EF", "SYS", c) Then
        EF_SendFile = ""
    Else
        EF_SendFile = Trim(c)
    End If

'---------------------------------------------- '対WELCAT　ﾗﾍﾞﾙ発行（枚数指定）用データファイル名取り込み 2015.10.06
    If GetIni(App.EXEName, "EI", "SYS", c) Then
        EI_SendFile = ""
    Else
        EI_SendFile = Trim(c)
    End If


'---------------------------------------------- '対WELCAT　ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙ発行用データファイル名(1)取り込み 2017.04.10
    If GetIni(App.EXEName, "R1", "SYS", c) Then
        R1_SendFile = ""
    Else
        R1_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙ発行用データファイル名(2)取り込み 2017.04.10
    If GetIni(App.EXEName, "R2", "SYS", c) Then
        R2_SendFile = ""
    Else
        R2_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙ発行用データファイル名(3)取り込み 2017.04.10
    If GetIni(App.EXEName, "R3", "SYS", c) Then
        R3_SendFile = ""
    Else
        R3_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙ発行用データファイル名(4)取り込み 2017.04.10
    If GetIni(App.EXEName, "R4", "SYS", c) Then
        R4_SendFile = ""
    Else
        R4_SendFile = Trim(c)
    End If





'---------------------------------------------- '出荷ログファイル名取り込み 2007.11.02
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

'---------------------------------------------- '共通メニュー情報取り込み
    If GetIni(App.EXEName, "ALL_MENU_GRP", "SYS", c) Then
        Beep
        MsgBox "共通メニュー情報の獲得に失敗しました。処理を中止します。"
        End
    End If


    ALL_MENU_GRP = Trim(c)

'---------------------------------------------- '検品チェック
    If GetIni(App.EXEName, "Inspection", "SYS", c) Then
        Inspection_Flg = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            Inspection_Flg = 1
        Else
            Inspection_Flg = CInt(Trim(c))
        End If
    End If
'---------------------------------------------- '検品数表示 2007.05.15
    If GetIni(App.EXEName, "Inspection_QTY", "SYS", c) Then
        Inspection_QTY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            Inspection_QTY = 0
        Else
            Inspection_QTY = CInt(Trim(c))
        End If
    End If

'---------------------------------------------- '出庫数入力　必須有無 2007.08.02
    If GetIni(App.EXEName, "SYUKO_QTY", "SYS", c) Then
        SYUKO_QTY = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            SYUKO_QTY = 0
        Else
            SYUKO_QTY = CInt(Trim(c))
        End If
    End If


'---------------------------------------------- '再検品チェック 2007.10.10
    If GetIni(App.EXEName, "Inspection_CHK", "SYS", c) Then
        Inspection_CHK = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            Inspection_CHK = 0
        Else
            Inspection_CHK = CInt(Trim(c))
        End If
    End If

'---------------------------------------------- '在庫照合メモ項目
    If GetIni(App.EXEName, "B2_MEMO", "SYS", c) Then
        B2_MEMO = ""
    Else
        B2_MEMO = Trim(c)
    End If
'--
    If GetIni(App.EXEName, "B8_MEMO", "SYS", c) Then
        B8_MEMO = ""
    Else
        B8_MEMO = Trim(c)
    End If
'---------------------------------------------- 'ファイルリトライ回数取り込み
    If GetIni("SYSTEM", "RETRY", "SYS", c) Then
        FILE_RETRY = 1
    Else
        If Not IsNumeric(Trim(c)) Then
            FILE_RETRY = 1
        Else
            FILE_RETRY = CInt(Trim(c))
        End If
    End If


'---------------------------------------------- '大阪ＰＣ運送会社読み替え用 2006.12.22
'    If GetIni(App.EXEName, "UNSOU_KAISHA", "SYS", c) Then
'        UNSOU_KAISHA_CODE = ""
'        UNSOU_KAISHA_NAME = ""
'    Else
'        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
'
'        If UBound(wkUNSOU_KAISHA) > 0 Then
'
'            UNSOU_KAISHA_CODE = wkUNSOU_KAISHA(0)
'            UNSOU_KAISHA_NAME = wkUNSOU_KAISHA(1)
'
'        Else
'            UNSOU_KAISHA_CODE = ""
'            UNSOU_KAISHA_NAME = ""
'        End If
'    End If
'
'
'---------------------------------------------- '大阪ＰＣ（新）運送会社読み替え用 2007.01.09
'
''久留米
'    KURUME_F = False
'    If GetIni(App.EXEName, "KURUME", "SYS", c) Then
'    Else
'        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
'        For i = 0 To UBound(wkUNSOU_KAISHA)
'            ReDim Preserve KURUME(0 To i)
'            KURUME(i) = Trim(wkUNSOU_KAISHA(i))
'            KURUME_F = True
'
'        Next i
'    End If
'
''    If UBound(KURUME) > 0 Then
''        KURUME_F = True
''    End If
''福山
'    FUKUYAMA_F = False
'    If GetIni(App.EXEName, "FUKUYAMA", "SYS", c) Then
'    Else
'        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
'        For i = 0 To UBound(wkUNSOU_KAISHA)
'            ReDim Preserve FUKUYAMA(0 To i)
'            FUKUYAMA(i) = Trim(wkUNSOU_KAISHA(i))
'            FUKUYAMA_F = True
'        Next i
'    End If
''    If UBound(FUKUYAMA) > 0 Then
''        FUKUYAMA_F = True
''    End If
''佐川
'    SAGAWA_F = False
'    If GetIni(App.EXEName, "SAGAWA", "SYS", c) Then
'    Else
'        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
'        For i = 0 To UBound(wkUNSOU_KAISHA)
'            ReDim Preserve SAGAWA(0 To i)
'            SAGAWA(i) = Trim(wkUNSOU_KAISHA(i))
'            SAGAWA_F = True
'        Next i
'    End If
'
''    If UBound(SAGAWA) > 0 Then
''        SAGAWA_F = True
''    End If





'福山
    If GetIni("FUKUYAMA", "length", "UNSOU", c) Then
        FUKUYAMA_Name = ""
        FUKUYAMA_Length = 0
        ReDim Preserve FUKUYAMA_CODE(0 To 0)
        FUKUYAMA_CODE(0) = "*"
    Else
        FUKUYAMA_Length = Val(Trim(c))
        
        If GetIni("FUKUYAMA", "Name", "UNSOU", c) Then
            FUKUYAMA_Name = ""
        Else
            FUKUYAMA_Name = Trim(c)
        End If
        
        If GetIni("FUKUYAMA", "Code", "UNSOU", c) Then
            ReDim Preserve FUKUYAMA_CODE(0 To 0)
            FUKUYAMA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve FUKUYAMA_CODE(0 To i)
                FUKUYAMA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If

'西武
    If GetIni("SEIBU", "length", "UNSOU", c) Then
        SEIBU_Name = ""
        SEIBU_Length = 0
        ReDim Preserve SEIBU_CODE(0 To 0)
        SEIBU_CODE(0) = "*"
    Else
        SEIBU_Length = Val(Trim(c))
        
        If GetIni("SEIBU", "Name", "UNSOU", c) Then
            SEIBU_Name = ""
        Else
            SEIBU_Name = Trim(c)
        End If
        
        If GetIni("SEIBU", "Code", "UNSOU", c) Then
            ReDim Preserve SEIBU_CODE(0 To 0)
            SEIBU_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SEIBU_CODE(0 To i)
                SEIBU_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If

'久留米
    If GetIni("KURUME", "length", "UNSOU", c) Then
        KURUME_Name = ""
        KURUME_Length = 0
        ReDim Preserve KURUME_CODE(0 To 0)
        KURUME_CODE(0) = "*"
    Else
        KURUME_Length = Val(Trim(c))
        If GetIni("KURUME", "Name", "UNSOU", c) Then
            KURUME_Name = ""
        Else
            KURUME_Name = Trim(c)
        End If
        If GetIni("KURUME", "Code", "UNSOU", c) Then
            ReDim Preserve KURUME_CODE(0 To 0)
            KURUME_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve KURUME_CODE(0 To i)
                KURUME_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If

'佐川
    If GetIni("SAGAWA", "length", "UNSOU", c) Then
        SAGAWA_Name = ""
        SAGAWA_Length = 0
        ReDim Preserve SAGAWA_CODE(0 To 0)
        SAGAWA_CODE(0) = "*"
    Else
        SAGAWA_Length = Val(Trim(c))
        
        If GetIni("SAGAWA", "Name", "UNSOU", c) Then
            SAGAWA_Name = ""
        Else
            SAGAWA_Name = Trim(c)
        End If
        
        If GetIni("SAGAWA", "Code", "UNSOU", c) Then
            ReDim Preserve SEIBU_CODE(0 To 0)
            SAGAWA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SAGAWA_CODE(0 To i)
                SAGAWA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If

'ヤマト
    If GetIni("YAMATO", "length", "UNSOU", c) Then
        YAMATO_Name = ""
        YAMATO_Length = 0
        ReDim Preserve YAMATO_CODE(0 To 0)
        YAMATO_CODE(0) = "*"
    Else
        YAMATO_Length = Val(Trim(c))
        If GetIni("YAMATO", "Name", "UNSOU", c) Then
            YAMATO_Name = ""
        Else
            YAMATO_Name = Trim(c)
        End If
        If GetIni("YAMATO", "Code", "UNSOU", c) Then
            ReDim Preserve YAMATO_CODE(0 To 0)
            YAMATO_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve YAMATO_CODE(0 To i)
                YAMATO_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If
'ﾊﾟﾅﾎｰﾑ 2011.06.06
    If GetIni("PANAHOME", "length", "UNSOU", c) Then
        PANA_Name = ""
        PANA_Length = 0
        ReDim Preserve PANA_CODE(0 To 0)
        PANA_CODE(0) = "*"
    Else
        PANA_Length = Val(Trim(c))
        
        If GetIni("PANAHOME", "Name", "UNSOU", c) Then
            PANA_Name = ""
        Else
            PANA_Name = Trim(c)
        End If
        If GetIni("PANAHOME", "Code", "UNSOU", c) Then
            ReDim Preserve PANA_CODE(0 To 0)
            PANA_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve PANA_CODE(0 To i)
                PANA_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If
    
'積水 2011.06.06
    If GetIni("SEKISUI", "length", "UNSOU", c) Then
        SEKISUI_Name = ""
        SEKISUI_Length = 0
        ReDim Preserve SEKISUI_CODE(0 To 0)
        SEKISUI_CODE(0) = "*"
    Else
        SEKISUI_Length = Val(Trim(c))
        If GetIni("SEKISUI", "Name", "UNSOU", c) Then
            SEKISUI_Name = ""
        Else
            SEKISUI_Name = Trim(c)
        End If
        If GetIni("SEKISUI", "Code", "UNSOU", c) Then
            ReDim Preserve SEKISUI_CODE(0 To 0)
            SEKISUI_CODE(0) = "*"
        Else
            wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
            For i = 0 To UBound(wkUNSOU_KAISHA)
                ReDim Preserve SEKISUI_CODE(0 To i)
                SEKISUI_CODE(i) = wkUNSOU_KAISHA(i)
            Next i
        End If
    End If

'---------------------------------------------- 'チャーター便ｺｰﾄﾞ   2010.01.21
    If GetIni("CHARTER", "Code", "UNSOU", c) Then
        KEN_CHARTER_CD = "*"
    Else
        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        KEN_CHARTER_CD = wkUNSOU_KAISHA(0)
        If GetIni("CHARTER", "Name", "UNSOU", c) Then
            KEN_CHARTER_NM = ""
        Else
            KEN_CHARTER_NM = Trim(c)
        End If
    End If
'---------------------------------------------- '赤帽便ｺｰﾄﾞ   2010.01.21
    If GetIni("AKABOU", "Code", "UNSOU", c) Then
        KEN_AKABOU_CD = "*"
    Else
        wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
        KEN_AKABOU_CD = wkUNSOU_KAISHA(0)
        If GetIni("AKABOU", "Name", "UNSOU", c) Then
            KEN_AKABOU_NM = ""
        Else
            KEN_AKABOU_NM = Trim(c)
        End If
    End If
'---------------------------------------------- 'ロジステックｺｰﾄﾞ   2010.01.21
    If GetIni(App.EXEName, "KEN_LOGISTIC", "SYS", c) Then
        KEN_LOGISTIC_CD = "*"
    Else
        KEN_LOGISTIC_CD = Trim(c)
    End If
'---------------------------------------------- '運送会社取込み   2018.10.05
    Erase UNSOU_TBL
    
    If GetIni(App.EXEName, "UNSOU_SELECT", "SYS", c) Then
        c = "運送会社未登録"
    Else
    End If
    wkUNSOU_KAISHA = Split(Trim(c), ",", -1)
    For i = 0 To UBound(wkUNSOU_KAISHA)
        ReDim Preserve UNSOU_TBL(0 To i)
        UNSOU_TBL(i).CODE = i + 1
        UNSOU_TBL(i).NAME = wkUNSOU_KAISHA(i)
    Next i
'---------------------------------------------- '送り先集約CD       2015.04.27
    Erase COL_OKURISAKI_CD
    Erase COL_OKURISAKI_NAME
    i = 1
    Do
        If GetIni("送り先集約CD指定", Format(i, "00"), "UNSOU", c) Then
            If i = 1 Then
                ReDim COL_OKURISAKI_CD(0 To 0)
                ReDim COL_OKURISAKI_NAME(0 To 0)
                COL_OKURISAKI_CD(0) = "********************"
                COL_OKURISAKI_NAME(0) = "********************"
            End If
            Exit Do
        End If
        wkVariant = Split(Trim(c), ",", -1)
        ReDim Preserve COL_OKURISAKI_CD(0 To i - 1)
        ReDim Preserve COL_OKURISAKI_NAME(0 To i - 1)
        
        COL_OKURISAKI_CD(i - 1) = CStr(wkVariant(0))
        COL_OKURISAKI_NAME(i - 1) = CStr(wkVariant(1))

        i = i + 1
    Loop
'---------------------------------------------- '送り先CD           2015.04.27
    Erase OKURISAKI_CD
    Erase OKURISAKI_NAME

    i = 1
    Do
        If GetIni("送り先CD指定", Format(i, "00"), "UNSOU", c) Then
            
            If i = 1 Then
            
                ReDim OKURISAKI_CD(0 To 0)
                ReDim OKURISAKI_NAME(0 To 0)
                
                OKURISAKI_CD(0) = "********************"
                OKURISAKI_NAME(0) = "********************"
            End If
            Exit Do
        End If

        wkVariant = Split(Trim(c), ",", -1)
        ReDim Preserve OKURISAKI_CD(0 To i - 1)
        ReDim Preserve OKURISAKI_NAME(0 To i - 1)
            
        OKURISAKI_CD(i - 1) = CStr(wkVariant(0))
        OKURISAKI_NAME(i - 1) = CStr(wkVariant(1))
        i = i + 1
    Loop

'---------------------------------------------- '事業部読み替え有無 2008.07.24
    If GetIni(App.EXEName, "JGYOBU_YOMIKAE", "SYS", c) Then
        JGYOBU_YOMIKAE_F = False
    Else
        JGYOBU_YOMIKAE_F = True
    
        JGYOBU_YOMIKAE_T = Split(Trim(c), ",", -1)
    
    End If

'---------------------------------------------- '担当者照合 2008.08.08
    If GetIni(App.EXEName, "TANTO_REF", "SYS", c) Then
        TANTO_REF = ""
    Else
        TANTO_REF = Trim(c)
    End If
'---------------------------------------------- 'メニュー終了 2008.08.08
    If GetIni(App.EXEName, "MENU_FIN", "SYS", c) Then
        MENU_FIN = ""
    Else
        MENU_FIN = Trim(c)
    End If
'---------------------------------------------- 'ｷｬﾝｾﾙ時の動作 2008.09.01
    CANCEL_OPE = True
    If GetIni(App.EXEName, "CANCEL_OPE", "SYS", c) Then
    Else
        
        If Trim(c) = "1" Then
            CANCEL_OPE = False
        End If
    End If

'---------------------------------------------- '在庫精査の要因 2008.11.20
    ZAIKO_SEISA_PURA = "11"
    
    If GetIni(App.EXEName, "ZAIKO_SEISA_PURA", "SYS", c) Then
    Else
        ZAIKO_SEISA_PURA = Trim(c)
    End If


    ZAIKO_SEISA_MAINA = "21"
    
    If GetIni(App.EXEName, "ZAIKO_SEISA_MAINA", "SYS", c) Then
    Else
        ZAIKO_SEISA_MAINA = Trim(c)
    End If



'---------------------------------------------- '出荷検品時のBUZZER音 2009.07.27
    If GetIni(App.EXEName, "Inspe_BUZZER", "SYS", c) Then
    
        Wel_Inspe_BUZZER = Buzzer_DEF
    Else
        Wel_Inspe_BUZZER = Trim(c)
        
    End If


'---------------------------------------------- '出荷検品時のBUZZER音 2009.07.27
    If GetIni(App.EXEName, "SYUKA_BUZZER", "SYS", c) Then
    
        Wel_SYUKA_BUZZER = Buzzer_DEF
    Else
        Wel_SYUKA_BUZZER = Trim(c)
        
    End If


'---------------------------------------------- '口数入力可／不可 2010.02.17
    If GetIni(App.EXEName, "Inspection_Input", "SYS", c) Then
    
        Inspection_Input = False
    Else
        
        
        If Trim(c) = "1" Then
            Inspection_Input = True
        Else
            Inspection_Input = False
        End If
        
    End If


'---------------------------------------------- '廃棄用向け先 2010.02.22
    HAIKI_MTS_F = False
    If GetIni(App.EXEName, "HAIKI_MTS", "SYS", c) Then
    
    Else
        
        HAIKI_MTS = Split(Trim(c), ",", -1)
        HAIKI_MTS_F = True
        
    End If
'---------------------------------------------- '異常対処 2011.04.02
    If GetIni(App.EXEName, "RECOVER_F", "SYS", c) Then
        RECOVER_F = False
    Else
        If Trim(c) = "1" Then
            RECOVER_F = True
        Else
            RECOVER_F = False
        End If
    End If


'---------------------------------------------- '構成チェック用種別 2011.04.18
    If GetIni(App.EXEName, "kousei_check", "SYS", c) Then
        Kousei_check_F = False
    Else
    
        Kousei_check_Tb = Split(Trim(c), ",", -1)
        Kousei_check_F = True
    
    
    End If


'---------------------------------------------- '大阪　積水向け処理 2011.05.27
    
    '集合梱包済み
    If GetIni(App.EXEName, "SEK_KONPO_F", "SYS", c) Then
        SEK_KONPO_F = False
    Else
        If Trim(c) = "1" Then
            SEK_KONPO_F = True
        Else
            SEK_KONPO_F = False
        End If
    End If

    '照合済み
    If GetIni(App.EXEName, "SEK_KEN_SHOGO_F", "SYS", c) Then
        SEK_KEN_SHOGO_F = False
    Else
        If Trim(c) = "1" Then
            SEK_KEN_SHOGO_F = True
        Else
            SEK_KEN_SHOGO_F = False
        End If
    End If

    '梱包済み
    If GetIni(App.EXEName, "SEK_KEN_KONPO_F", "SYS", c) Then
        SEK_KEN_KONPO_F = False
    Else
        If Trim(c) = "1" Then
            SEK_KEN_KONPO_F = True
        Else
            SEK_KEN_KONPO_F = False
        End If
    End If


    '検品済み
    If GetIni(App.EXEName, "SEK_KEN_KENPIN_F", "SYS", c) Then
        SEK_KEN_KENPIN_F = True
    Else
        If Trim(c) = "1" Then
            SEK_KEN_KENPIN_F = False
        Else
            SEK_KEN_KENPIN_F = True
        End If
    End If


    '検品時ﾗﾍﾞﾙ開始ﾍﾟｰｼﾞ
    If GetIni(App.EXEName, "SEK_LABEL_PAGE", "SYS", c) Then
        SEK_LABEL_PAGE = 0
    Else
        If IsNumeric(Trim(c)) Then
            SEK_LABEL_PAGE = Val(Trim(c))
        Else
            SEK_LABEL_PAGE = 0
        End If
    End If


'---------------------------------------------- '大阪　積水向け処理 2011.05.27

'---------------------------------------------- '大阪　積水向け移動出庫要因 2011.06.15
    If GetIni(App.EXEName, "SEK_IDO_SYUKO", "SYS", c) Then
        WEL_SEK_IDO_SYUKO = "**"
    Else
        WEL_SEK_IDO_SYUKO = Trim(c)
    End If
'---------------------------------------------- '大阪　積水向け移動出庫要因 2011.06.15





'---------------------------------------------- '月平均／生産計画用 2011.07.05
    If GetIni("F120050", "TUKI1", "F120050", c) Then
        TUKI1 = 3
    Else
        If IsNumeric(Trim(c)) Then
            TUKI1 = Val(Trim(c))
        Else
            TUKI1 = 3
        End If
    End If

    If GetIni("F120050", "TUKI2", "F120050", c) Then
        TUKI2 = 3
    Else
        If IsNumeric(Trim(c)) Then
            TUKI2 = Val(Trim(c))
        Else
            TUKI2 = 3
        End If
    End If
'---------------------------------------------- '大阪PC 部材関係 2012.03.06
    If GetIni("F110010", "SHIMEBI", "SYS", c) Then
        SHIMEBI = "25"
    Else
        If IsNumeric(Trim(c)) Then
            SHIMEBI = Trim(c)
        Else
            SHIMEBI = "25"
        End If
    End If

    '開始日
    If Mid(Format(Date, "YYYYMMDD"), 7, 2) > SHIMEBI Then
        BUZAI_DATE_S = Mid(Format(Date, "YYYYMMDD"), 1, 6) & Format(Val(SHIMEBI) + 1, "00")
    Else
        wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
        wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) - 1
        If wkMM < 1 Then
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) - 1
            wkMM = 12
        End If
        BUZAI_DATE_S = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(SHIMEBI) + 1, "00")
    End If
    '終了日
    If Mid(Format(Date, "YYYYMMDD"), 7, 2) <= SHIMEBI Then
        BUZAI_DATE_E = Mid(Format(Date, "YYYYMMDD"), 1, 6) & SHIMEBI
    Else
        wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4))
        wkMM = Val(Mid(Format(Date, "YYYYMMDD"), 5, 2)) + 1
        If wkMM > 12 Then
            wkYY = Val(Mid(Format(Date, "YYYYMMDD"), 1, 4)) + 1
            wkMM = 1
        End If
        BUZAI_DATE_E = Format(wkYY, "0000") & Format(wkMM, "00") & Format(Val(SHIMEBI) + 1, "00")
    End If
    '使用月
    USE_YM = Mid(BUZAI_DATE_E, 1, 6)


    '検収～棚入れの要因
    If GetIni("F110010", "IN_TANA_S_OSAKA", "SYS", c) Then
        c = "**"
        IN_TANA_S_OSAKA = Split(Trim(c), ",", -1)
    Else
        IN_TANA_S_OSAKA = Split(Trim(c), ",", -1)
    End If




    '出庫処理
    If GetIni("F110010", "IDO_OUT_OSAKA", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材出庫処理の要因を登録してください SYS.INI [F110010] IDO_OUT_OSAKA="
        Unload Me
    Else
        Wel_IDO_OUT_OSAKA = Trim(c)
    End If

    '出庫処理2   2014.11.07
    If GetIni("F110010", "IDO_OUT_OSAKA2", "SYS", c) Then
        Call LOG_OUT(LOG_F, "大阪ＰＣ　部材出庫処理2の要因を登録してください SYS.INI [F110010] IDO_OUT_OSAKA2=")
        c = "**"
    End If
    Wel_IDO_OUT_OSAKA2 = Trim(c)


    '出庫処理3   2016.05.11
    If GetIni("F110010", "IDO_OUT_OSAKA3", "SYS", c) Then
        Call LOG_OUT(LOG_F, "大阪ＰＣ　部材出庫処理2の要因を登録してください SYS.INI [F110010] IDO_OUT_OSAKA3=")
        c = "**"
    End If
    Wel_IDO_OUT_OSAKA3 = Trim(c)




    '振替出庫／入庫
    If GetIni("F110010", "HIN_FURIKA_S", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材振替出庫処理の要因を登録してください SYS.INI [F110010] HIN_FURIKA_S="
        Unload Me
    Else
        Wel_HIN_FURIKA_S = Trim(c)
    End If

    '振替出庫
    If GetIni("F110010", "BEF_HINBAN_S", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材振替出庫処理の要因を登録してください SYS.INI [F110010] BEF_HINBAN_S="
        Unload Me
    Else
        YOIN_BEF_HINBAN_S = Trim(c)
    End If
    
    '振替入庫
    If GetIni("F110010", "AFT_HINBAN_S", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材振替入庫処理の要因を登録してください SYS.INI [F110010] AFT_HINBAN_S="
        Unload Me
    Else
        YOIN_AFT_HINBAN_S = Trim(c)
    End If

    '振替出庫／入庫
    If GetIni("F110010", "IDO_IN_OSAKA", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材振替入庫処理の要因処理の要因を登録してください SYS.INI [F110010] IDO_IN_OSAKA="
        Unload Me
    Else
        Wel_IDO_IN_OSAKA = Trim(c)
    End If


    '構成チェック       2012.03.16
    If GetIni("F110010", "COMPO_OSAKA_CHECK", "SYS", c) Then
        MsgBox "大阪ＰＣ　部材構成チェック処理の要因を登録してください SYS.INI [F110010] COMPO_OSAKA_CHECK="
        Unload Me
    Else
        Wel_COMPO_OSAKA_CHECK = Trim(c)
    End If


    '検品               2012.03.18
    If GetIni("F110010", "KENPIN_OSAKA", "SYS", c) Then
        MsgBox "大阪ＰＣ　検品処理の要因を登録してください SYS.INI [F110010] KENPIN_OSAKA="
        Unload Me
    Else
        Wel_KENPIN_OSAKA = Trim(c)
    End If
    '検品時引落倉庫     2012.03.18
    If GetIni("F110010", "KENPIN_SOKO", "SYS", c) Then
        MsgBox "大阪ＰＣ　検品処理時引落倉庫を登録してください SYS.INI [F110010] KENPIN_SOKO="
        Unload Me
    Else
        KENPIN_OSAKA_SOKO_No = Trim(c)
    End If
    
    '検品               2016.05.20
    If GetIni("F110010", "KENPIN_OSAKA_NEW", "SYS", c) Then
'        MsgBox "大阪ＰＣ　検品処理の要因を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW="
        Call LOG_OUT(LOG_F, "大阪ＰＣ　検品処理の要因を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW=")
        Wel_KENPIN_OSAKA_NEW = "**"
    Else
        Wel_KENPIN_OSAKA_NEW = Trim(c)
    End If
    
    
    '検品2               2016.06.27
    If GetIni("F110010", "KENPIN_OSAKA_NEW2", "SYS", c) Then
'        MsgBox "大阪ＰＣ　検品処理の要因を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW="
        Call LOG_OUT(LOG_F, "大阪ＰＣ　検品処理の要因(指図票完了)を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW2=")
        Wel_KENPIN_OSAKA_NEW2 = "**"
    Else
        Wel_KENPIN_OSAKA_NEW2 = Trim(c)
    End If
    
    '検品2               2016.06.27
    If GetIni("F110010", "KENPIN_OSAKA_NEW2_FILE", "SYS", c) Then
'        MsgBox "大阪ＰＣ　検品処理の要因を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW="
        Call LOG_OUT(LOG_F, "大阪ＰＣ　検品処理の要因(指図票完了)を登録してください SYS.INI [F110010] KENPIN_OSAKA_NEW2_FILE=")
        KENPIN_OSAKA_NEW2_FILE = ""
    Else
        KENPIN_OSAKA_NEW2_FILE = Trim(c)
    End If
    
    
    '商品化完了登録               2018.04.26
'    If GetIni("F110010", "SHIJI_END", "SYS", c) Then
'        Call LOG_OUT(LOG_F, "指図票完了登録の要因を登録してください SYS.INI [F110010] SHIJI_END=")
'        Wel_SHIJI_END = ""
'    Else
'        Wel_SHIJI_END = Trim(c)
'    End If
    '商品化完了登録               2018.04.26
    If GetIni("ACTION", "ACTION_WEL_PARA2309", "SYS", c) Then
        Call LOG_OUT(LOG_F, "指図票完了登録の要因を登録してください SYS.INI [ACTION] ACTION_WEL_PARA2309=")
        LCD_SHIJI = "指示№"
    Else
        
        wkVariant = Split(Trim(c), ",", -1)
        
        
        LCD_SHIJI = wkVariant(3)
    End If
    WEL_Para_Tbl(22, 8).Wel_Para(1).LCD = LCD_SHIJI
    
    
    
    
    '対WELCAT　ﾗﾍﾞﾙ発行用データファイル名取り込み
    If GetIni(App.EXEName, "88", "SYS", c) Then
        LABEL_88_SendFile = ""
    Else
        LABEL_88_SendFile = Trim(c)
    End If


    '部材出庫処理時の在庫集計対象倉庫   2014.10.29
    If GetIni(App.EXEName, "IDO_OUT_ZAIKO_SOKO", "SYS", c) Then
        IDO_OUT_ZAIKO_SOKO_F = False
    Else
        IDO_OUT_ZAIKO_SOKO_F = True
        Zaiko_Syukei_Jyogai_Soko_No2 = Split(Trim(c), ",", -1)
    End If

'---------------------------------------------- '大阪PC 部材関係 2012.03.06


'---------------------------------------------- '開始／終了 2012.11.06
    If GetIni("F110010", "STARTTIME", "SYS", c) Then
        wkTH = ""
    Else
        wkTH = Trim(c)
        If Not IsNumeric(wkTH) Then
            wkTH = ""
        End If
    End If
    Text2(2) = Mid(wkTH, 1, 2)
    Text2(3) = Mid(wkTH, 3, 2)


    If GetIni("F110010", "ENDTIME", "SYS", c) Then
        wkTH = ""
    Else
        wkTH = Trim(c)
        If Not IsNumeric(wkTH) Then
            wkTH = ""
        End If
    End If
    Text2(0) = Mid(wkTH, 1, 2)
    Text2(1) = Mid(wkTH, 3, 2)




'---------------------------------------------- '開始／終了 2012.11.06


'---------------------------------------------- '受信タイムアウト   2014.01.04
    If GetIni("F110010", "WAIT_TIME", "SYS", c) Then
        WAIT_TIME = 300
    Else
        If Not IsNumeric(Trim(c)) Then
            WAIT_TIME = 300
        Else
            WAIT_TIME = Val(c)
        End If
    End If



'---------------------------------------------- '製造№管理 2013.06.06
    
    
'--------   送信ファイル名
    If GetIni(App.EXEName, "N4", "SYS", c) Then
    
        N4_SendFile = ""
    Else
        N4_SendFile = Trim(c)
        
    End If
'--------   スキャナ名称    2014.07.01
    If GetIni("ACTION", "ACTION_WEL_PARA2401", "SYS", c) Then
        LCD_BCR_N1 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N1 = wkVariant(3)
        
            LCD_LotNo_BCR = wkVariant(3)
        Else
            LCD_BCR_N1 = ""
        End If
    End If


    If GetIni("ACTION", "ACTION_WEL_PARA2402", "SYS", c) Then
        LCD_BCR_N2 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N2 = wkVariant(3)
        Else
            LCD_BCR_N2 = ""
        End If
    End If

    If GetIni("ACTION", "ACTION_WEL_PARA2403", "SYS", c) Then
        LCD_BCR_N3 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N3 = wkVariant(3)
        Else
            LCD_BCR_N3 = ""
        End If
    End If

    If GetIni("ACTION", "ACTION_WEL_PARA2404", "SYS", c) Then
        LCD_BCR_N4 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N4 = wkVariant(3)
        Else
            LCD_BCR_N4 = ""
        End If
    End If

    If GetIni("ACTION", "ACTION_WEL_PARA2405", "SYS", c) Then
        LCD_BCR_N5 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N5 = wkVariant(3)
            LCD_InvNo_BCR = wkVariant(3)
        Else
            LCD_BCR_N5 = ""
        End If
    End If

    If GetIni("ACTION", "ACTION_WEL_PARA2406", "SYS", c) Then
        LCD_BCR_N6 = ""
    Else
        wkVariant = Split(Trim(c), ",", -1)
        If UBound(wkVariant) > 2 Then
            LCD_BCR_N6 = wkVariant(3)
        Else
            LCD_BCR_N6 = ""
        End If
    End If


'--------   スキャナ名称    2014.07.01
'---------------------------------------------- '製造№管理 2013.06.06



'---------------------------------------------- '数量検品要因 2014.03.05
    WEL_KENPIN_GAI = "L1"
    
    If GetIni(App.EXEName, "WEL_KENPIN_GAI", "SYS", c) Then
    Else
        WEL_KENPIN_GAI = Trim(c)
    End If

    WEL_KENPIN_Su = "L2"
    
    If GetIni(App.EXEName, "WEL_KENPIN_SU", "SYS", c) Then
    Else
        WEL_KENPIN_Su = Trim(c)
    End If
'---------------------------------------------- '数量検品要因 2014.03.05

'---------------------------------------------- '送受信ﾛｸﾞ 2014.03.19
    If GetIni(App.EXEName, "F110010_LOG", "SYS", c) Then
        F110010_LOG = ""
    Else
        F110010_LOG = Trim(c)
    End If
'---------------------------------------------- '送受信ﾛｸﾞ 2014.03.19


'---------------------------------------------- '奈良モジュール　対象倉庫 2014.07.01
    If GetIni(App.EXEName, "MODULE_SOKO", "SYS", c) Then
        c = "**"
        Nara_Soko_T = Split(Trim(c), ",", -1)
    Else
        Nara_Soko_T = Split(Trim(c), ",", -1)
    End If


'---------------------------------------------- '奈良モジュール　対象倉庫 2014.07.01


'---------------------------------------------- '奈良モジュール　要因 2018.10.03
    If GetIni(App.EXEName, "MODULE_IN", "SYS", c) Then
        Wel_MODULE_IN = "**"
    Else
        Wel_MODULE_IN = Trim(c)
    End If


'---------------------------------------------- '奈良モジュール　要因 2018.10.03

'---------------------------------------------- '奈良モジュール　入庫倉庫 2018.10.03-->
    If GetIni(App.EXEName, "MODULE_IN_SOKO", "SYS", c) Then
        If Wel_MODULE_IN <> "**" Then

            MsgBox "奈良モジュール　入庫倉庫([F110010] [MODULE_IN_SOKO])の獲得に失敗しました。処理を中止します。"
            End
        End If
    Else
        Wel_MODULE_IN_SOKO = Trim(c)
    End If


'---------------------------------------------- '奈良モジュール　対象倉庫 2014.07.01



'---------------------------------------------- '画面タイトル 2014.07.01
    If GetIni(App.EXEName, "TITLE", "SYS", c) Then
        MAIN_TITLE = ""
    Else
        MAIN_TITLE = Trim(c)
    End If

    Label1.Caption = MAIN_TITLE & F1100101.Caption
'---------------------------------------------- '画面タイトル 2014.07.01

'---------------------------------------------- '特定送り先 2015.12.22
    If GetIni(App.EXEName, "KONPOU_OKURISAKI_CD", "SYS", c) Then
        c = "********"
    End If
    wkVariant = Split(Trim(c), ",", -1)

    Erase KONPOU_OKURISAKI_CD

    For i = 0 To UBound(wkVariant)
        ReDim Preserve KONPOU_OKURISAKI_CD(0 To i)
        KONPOU_OKURISAKI_CD(i) = Trim(wkVariant(i))
    Next i

'---------------------------------------------- '特定送り先 2015.12.22
'------------------------------------------------------------------ ラベルプリンター印字濃度    2016.02.15
    If GetIni(App.EXEName, "DK", "SYS", c) Then
        DK_DEF = 12
    Else
        If Not IsNumeric(Trim(c)) Then
            DK_DEF = 12
        Else
            If CInt(Trim(c)) < 1 Or CInt(Trim(c)) > 16 Then
                DK_DEF = 12
            Else
                DK_DEF = CInt(c)
            End If
        End If
    End If
'------------------------------------------------------------------ ラベルプリンター印字濃度    2016.02.15

'------------------------------------------------------------------ ラベルプリンター印字 設定   2017.04.14
    If GetIni("LABEL", "LABEL01_DEF", "F110010LABEL", c) Then
        LABEL01_DEF = ""
    Else
        LABEL01_DEF = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_HIN_F", "F110010LABEL", c) Then
        LABEL01_HIN_F = ""
    Else
        LABEL01_HIN_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_HIN_T", "F110010LABEL", c) Then
        LABEL01_HIN_T = ""
    Else
        LABEL01_HIN_T = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_HIN_B", "F110010LABEL", c) Then
        LABEL01_HIN_B = ""
    Else
        LABEL01_HIN_B = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_BIK_F", "F110010LABEL", c) Then
        LABEL01_BIK_F = ""
    Else
        LABEL01_BIK_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_BIK_T", "F110010LABEL", c) Then
        LABEL01_BIK_T = ""
    Else
        LABEL01_BIK_T = Trim(c)
    End If
    
    
    If GetIni("LABEL", "LABEL01_IRI_F", "F110010LABEL", c) Then
        LABEL01_IRI_F = ""
    Else
        LABEL01_IRI_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_IRI_T", "F110010LABEL", c) Then
        LABEL01_IRI_T = ""
    Else
        LABEL01_IRI_T = Trim(c)
    End If
    
    
    
    
    
    If GetIni("LABEL", "LABEL01_LOC_F", "F110010LABEL", c) Then
        LABEL01_LOC_F = ""
    Else
        LABEL01_LOC_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_LOC_T", "F110010LABEL", c) Then
        LABEL01_LOC_T = ""
    Else
        LABEL01_LOC_T = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_LOC_B", "F110010LABEL", c) Then
        LABEL01_LOC_B = ""
    Else
        LABEL01_LOC_B = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_QTY_F", "F110010LABEL", c) Then
        LABEL01_QTY_F = ""
    Else
        LABEL01_QTY_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL01_QTY_T", "F110010LABEL", c) Then
        LABEL01_QTY_T = ""
    Else
        LABEL01_QTY_T = Trim(c)
    End If

'---
    If GetIni("LABEL", "LABEL02_DEF", "F110010LABEL", c) Then
        LABEL02_DEF = ""
    Else
        LABEL02_DEF = Trim(c)
    End If
    If GetIni("LABEL", "LABEL02_HIN_F", "F110010LABEL", c) Then
        LABEL02_HIN_F = ""
    Else
        LABEL02_HIN_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL02_HIN_T", "F110010LABEL", c) Then
        LABEL02_HIN_T = ""
    Else
        LABEL02_HIN_T = Trim(c)
    End If
    If GetIni("LABEL", "LABEL02_HIN_B", "F110010LABEL", c) Then
        LABEL02_HIN_B = ""
    Else
        LABEL02_HIN_B = Trim(c)
    End If

'---
    If GetIni("LABEL", "LABEL03_DEF", "F110010LABEL", c) Then
        LABEL03_DEF = ""
    Else
        LABEL03_DEF = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_ID_F", "F110010LABEL", c) Then
        LABEL03_ID_F = ""
    Else
        LABEL03_ID_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_ID_T", "F110010LABEL", c) Then
        LABEL03_ID_T = ""
    Else
        LABEL03_ID_T = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_ID_B", "F110010LABEL", c) Then
        LABEL03_ID_B = ""
    Else
        LABEL03_ID_B = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_UN_F", "F110010LABEL", c) Then
        LABEL03_UN_F = ""
    Else
        LABEL03_UN_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_UN_T", "F110010LABEL", c) Then
        LABEL03_UN_T = ""
    Else
        LABEL03_UN_T = Trim(c)
    End If


    If GetIni("LABEL", "LABEL03_OKURI_F", "F110010LABEL", c) Then
        LABEL03_OKURI_F = ""
    Else
        LABEL03_OKURI_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_OKURI_T", "F110010LABEL", c) Then
        LABEL03_OKURI_T = ""
    Else
        LABEL03_OKURI_T = Trim(c)
    End If

    If GetIni("LABEL", "LABEL03_DEN_F", "F110010LABEL", c) Then
        LABEL03_DEN_F = ""
    Else
        LABEL03_DEN_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL03_DEN_T", "F110010LABEL", c) Then
        LABEL03_DEN_T = ""
    Else
        LABEL03_DEN_T = Trim(c)
    End If

'---
    If GetIni("LABEL", "LABEL04_DEF", "F110010LABEL", c) Then
        LABEL04_DEF = ""
    Else
        LABEL04_DEF = Trim(c)
    End If
    If GetIni("LABEL", "LABEL04_LOC_F", "F110010LABEL", c) Then
        LABEL04_LOC_F = ""
    Else
        LABEL04_LOC_F = Trim(c)
    End If
    If GetIni("LABEL", "LABEL04_LOC_T", "F110010LABEL", c) Then
        LABEL04_LOC_T = ""
    Else
        LABEL04_LOC_T = Trim(c)
    End If
    If GetIni("LABEL", "LABEL04_LOC_B", "F110010LABEL", c) Then
        LABEL04_LOC_B = ""
    Else
        LABEL04_LOC_B = Trim(c)
    End If


'------------------------------------------------------------------ ラベルプリンター印字 設定   2017.04.14






'------------------------------------------------------------------ 大阪移動出庫　１未満時の丸め 2016.07.16
    If GetIni(App.EXEName, "DK", "SYS", c) Then
        IDO_OUT_OSAKA_RND = 0
    Else
        If Trim(c) = "1" Then
            IDO_OUT_OSAKA_RND = 1
        Else
            IDO_OUT_OSAKA_RND = 0
        End If
    End If
'------------------------------------------------------------------ 大阪移動出庫　１未満時の丸め 2016.07.16



'------------------------------------------------------------------ 着店コード印字　有無 2017.04.07
    If GetIni(App.EXEName, "TYAKUTEN_PRINT", "SYS", c) Then
        TYAKUTEN_PRINT = 0
    Else
        If Trim(c) = "1" Then
            TYAKUTEN_PRINT = 1
        Else
            TYAKUTEN_PRINT = 0
        End If
    End If
'------------------------------------------------------------------ 出庫(P0)在庫残表示　   2017.12.05
    If GetIni(App.EXEName, "ZAIKO_DISP_FLG", "SYS", c) Then
        ZAIKO_DISP_FLG = 0
    Else
        If Trim(c) = "1" Then
            ZAIKO_DISP_FLG = 1
        Else
            ZAIKO_DISP_FLG = 0
        End If
    End If


'---------------------------------------------- '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


'---------------------------------------------- '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '品目マスタ(ワーク)ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '品目振替マスタＯＰＥＮ 2011.06.01
    If FURIKAE_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- 'メニュー管理マスタＯＰＥＮ
    If P_MENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- 'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '担当者別メニューＯＰＥＮ
    If P_TMENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データ（移動処理用）ＯＰＥＮ
    If wZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データ（商品／未商品切り替え用）ＯＰＥＮ
    If tmpZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '前借データＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '出荷予定データＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


''    sts = BTRV(BtOpGetFirst, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K8_Y_SYU, Len(K8_Y_SYU), 8)
''    Select Case sts
''        Case BtNoErr
''        Case BtErrIvldKey
''
''            If Y_SYU_Create_Index() Then
''                Normal_End = True
''                Unload Me
''            End If
''
''        Case Else
''            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
''            Normal_End = True
''            Unload Me
''    End Select



'---------------------------------------------- '出荷予定データＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '資材前借ﾃﾞｰﾀＯＰＥＮ
    If P_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '作業実績ﾛｸﾞＯＰＥＮ
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＯＰＥＮ 2006.12.07
    If Y_SYU_H_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '入荷予定(大阪PC向け)ＯＰＥＮ 2007.06.07
    If Y_NYU_O_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If

'---------------------------------------------- '管理マスタＯＰＥＮ 2007.06.28
    If P_KANRI_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '資材注文ﾃﾞｰﾀＯＰＥＮ 2007.06.28
    If P_SHORDER_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '受払先ﾏｽﾀＯＰＥＮ 2007.06.28
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '資材受入履歴ＯＰＥＮ 2007.06.28
    If P_SHUKEIRE_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If

'---------------------------------------------- '指図票データ（親）ＯＰＥＮ 2010.09.03
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


'---------------------------------------------- '指図票データ（子）ＯＰＥＮ 2011.03.02
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


'---------------------------------------------- '邸別注文データＯＰＥＮ 2011.04.25
    If Y_SYU_TEI_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If

'---------------------------------------------- '発注検討：親品番注文データＯＰＥＮ 2012.03.18
    If ODR_ORDER_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If

'---------------------------------------------- '商品化指図受入履歴データＯＰＥＮ 2012.03.18
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


'---------------------------------------------- '床暖管理データ　ＯＰＥＮ 2013.06.06
    If LOTNO_Open(BtOpenNomal) Then
'        Normal_End = True                  処理継続    2013.06.12
'        Unload Me
    End If


'---------------------------------------------- '床暖送状№データ　ＯＰＥＮ 2013.06.30
    If INVNO_Open(BtOpenNomal) Then
'        Normal_End = True                  処理継続    2013.06.12
'        Unload Me
    End If


'---------------------------------------------- '品目モジュール　ＯＰＥＮ 2014.06.24
    If M_ITEM_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If


'---------------------------------------------- '入荷予定 ＯＰＥＮ 2015.01.21
    If Y_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If

'---------------------------------------------- '送り状№ﾃﾞｰﾀ ＯＰＥＮ 2016.10.14
    Inspe_DelvNo_F = 0
    If HTDelvNo_Open(BtOpenRead) Then
        Inspe_DelvNo_F = 1
    End If

'---------------------------------------------- 'Id送り状№ﾃﾞｰﾀ ＯＰＥＮ 2016.10.14
    If Inspe_DelvNo_F = 0 Then
        If HTIdDelv_Open(BtOpenNomal) Then
            Normal_End = True
            Unload Me
        End If
    End If
'---------------------------------------------- '直送先Idﾃﾞｰﾀ ＯＰＥＮ 2016.10.14
    If Inspe_DelvNo_F = 0 Then
        If HTDrctId_Open(BtOpenRead) Then
            Normal_End = True
            Unload Me
        End If
    End If
'---------------------------------------------- 'メニュー機能チェック（個別 or 共通）
'    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
'    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
'    Select Case sts
'        Case BtNoErr
'            Menu_Type = 1           '共通メニューで運用
'        Case BtErrKeyNotFound
            Menu_Type = 2           '担当者別メニューで運用
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "担当者別メニュー")
'            Unload Me
'    End Select
    
    
'    ST_LOG_OUT_F = True '2008.08.08

    
    F1100101.Caption = F1100101.Caption & " " & LAST_UPDATE_DAY ' '2019/10/30 集合梱包AD-HEJP4N-Cを追加
    
    'F1100101.Height = 4050              '2012.10.02            2013.06.06 DELETE
    'F1100101.Width = 6720               '2012.10.02            2013.06.06 DELETE
    
    '▼[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
    '----- アプリケーション起動日時 -----
    gbl_StartApp = Now
    
    '----- ログファイル出力 -----
    If GetIni(SEC_LOG, KEY_LOGWRITE, "SYS", c) Then
        gbl_LogCfg.m_LogWrite = DEF_LOGWRITE
    Else
        gbl_LogCfg.m_LogWrite = IIf(CInt(Trim(c)) <> 0, True, False)
    End If
    
    '----- ログファイル出力パス -----
    If GetIni(SEC_LOG, KEY_LOGPATH, "SYS", c) Then
        gbl_LogCfg.m_LogPath = App.Path & "\Log\"
    Else
        gbl_LogCfg.m_LogPath = Trim(c)
    End If
    Call MkDirEx(gbl_LogCfg.m_LogPath)
    
    '----- ログファイル保存日数 -----
    If GetIni(SEC_LOG, KEY_LOGSAVE, "SYS", c) Then
        gbl_LogCfg.m_LogSave = DEF_LOGSAVE
    Else
        gbl_LogCfg.m_LogSave = CLng(Trim(c))
    End If
    
    '----- ログファイル名 -----
    gbl_LogCfg.m_LogFName = GetFullPath(gbl_LogCfg.m_LogPath, App.EXEName) & "_" & Format$(gbl_StartApp, "yyyymmdd") & ".log"
    
    '----- ローカルポート番号 -----
    If GetIni(SEC_SOCKET, KEY_LOCALPORT, "SYS", c) Then
        gbl_SockCfg.m_LocalPort = DEF_LOCALPORT
    Else
        gbl_SockCfg.m_LocalPort = Trim(c)
    End If
    lblINI(3).Caption = "   SocketPort:" & gbl_SockCfg.m_LocalPort    '2014.07.01
    
    '----- ログファイルの保存期間チェック -----
    If gbl_LogCfg.m_LogSave > 0 Then
        '----- アプリケーションログファイルチェック -----
        Call DeleteLogFile(App.EXEName, gbl_LogCfg.m_LogSave)
    End If
    '▲[2014/02/10 - M.MATSUYAMA 追加(Ver2.0.0)] ソケット通信用追加
    
    Show

    If Data_Clear_Proc(1, "") Then
        MsgBox "データ初期設定が出来ませんでした。"
        Unload Me
    End If


    If tmpZaiko_Clear_Proc() Then
        MsgBox "データ初期設定が出来ませんでした。"
        Unload Me
    End If

    Label1.Caption = MAIN_TITLE & "[停止中]"

    
'    Command1(0).Value = True                '2012.05.18

'    Timer1.Enabled = True                   '2012.05.18

    Command1(0).Value = True            '2013.04.16
    Timer1.Enabled = True               '2013.04.16


End Sub



Private Function Normal_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『通常テキスト受信』
'
'-------------------------------------------------------
    Normal_Proc = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Start                 '子機電源ON（ここには来ない）
        
        Case Step_TANTO_REQ             '担当者要求（ここには来ない）
        
        Case Step_TANTO_RES             '担当者回答
    
            MENU_UP_F = False   '2008.08.08
    
    
            If Tanto_Check_Proc(Sendbuf) Then
                Exit Function
            End If
    
        Case Step_JGYOBU_REQ            '事業部要求（ここには来ない）
                
        Case Step_JGYOBU_RES            '事業部回答
            
            MENU_UP_F = False   '2008.08.08
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                
        Case Step_NAIGAI_REQ            '国内外要求（ここには来ない）
                
        Case Step_NAIGAI_RES            '国内外回答
            
            MENU_UP_F = False   '2008.08.08
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                                        
            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ      '2006.07.14
                                        'メニュー要求（ここには来ない）
'2006.01.30        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
        Case Step_MENU1_REQ, Step_MENU2_REQ
                                        'メニュー回答
'2006.01.30        Case Step_MENU1_RES, Step_MENU2_RES, Step_MENU3_RES
        Case Step_MENU1_RES, Step_MENU2_RES
    
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
    
    End Select
    
    Normal_Proc = False

End Function

Private Sub Form_Unload(Cancel As Integer)

Dim sts As Integer


Dim yn  As Integer  '2013.06.06

    
    '2013.06.06 終了確認
    If Not Auto_Off Then
        yn = MsgBox("終了します。よろしいですか？", vbYesNo + vbDefaultButton2, "確認入力")
        If yn = vbNo Then
            Cancel = True
            Exit Sub
        Else
            CtrsWsk1.Unbind
            
            Normal_End = False              '正常終了
            Next_Step = 0                   '次処理起動しない
        End If
    End If
    '2013.06.06 終了確認
    



'---------------------------------------------- '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
'---------------------------------------------- '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
'---------------------------------------------- '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
'---------------------------------------------- '品目マスタ（ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
'---------------------------------------------- '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
'---------------------------------------------- '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
'---------------------------------------------- '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
'---------------------------------------------- '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
'---------------------------------------------- 'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
'---------------------------------------------- 'メニュー管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ")
        End If
    End If
'---------------------------------------------- '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
'---------------------------------------------- '担当者別メニューＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者別メニュー")
        End If
    End If
'---------------------------------------------- '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
'---------------------------------------------- '在庫データ（移動処理用）ＣＬＯＳＥ

    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
'---------------------------------------------- '前借データＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "前借データ")
        End If
    End If
'---------------------------------------------- '出荷予定データＣＬＯＳＥ
    
''    sts = BTRV(BtOpDropSupIndex, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K8_Y_SYU, Len(K8_Y_SYU), 8)
''    If sts Then
''        Call File_Error(sts, BtOpDropSupIndex, "出荷予定データ")
''    End If
    
    
    
    
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
'---------------------------------------------- '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
'---------------------------------------------- '在庫移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴データ")
        End If
    End If
'---------------------------------------------- '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
'---------------------------------------------- '月平均出荷数ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷数")
        End If
    End If
'---------------------------------------------- '資材前借ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材前借ﾃﾞｰﾀ")
        End If
    End If
'---------------------------------------------- '作業実績ﾛｸﾞＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材前借ﾃﾞｰﾀ")
        End If
    End If
'---------------------------------------------- '出荷予定(ﾎｽﾄｲﾒｰｼﾞ)ＣＬＯＳＥ   2006.12.07
    sts = BTRV(BtOpClose, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定(ﾎｽﾄｲﾒｰｼﾞ)データ")
        End If
    End If
'---------------------------------------------- '入荷予定(大阪PC向け)ＣＬＯＳＥ   2007.06.07
    sts = BTRV(BtOpClose, Y_NYU_O_POS, Y_NYU_O_REC, Len(Y_NYU_O_REC), K0_Y_NYU_O, Len(K0_Y_NYU_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定(大阪PC向け)データ")
        End If
    End If
'---------------------------------------------- 'ファイル環境リセット
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

'Call LOG_OUT(LOG_F, "END 1")

    If Next_Step = 1 Then
        sts = Shell("d:\newsdc\exe\F1100501.bat", vbNormalFocus)
        If sts = 0 Then
            MsgBox "[F110050]終了処理の起動に失敗しました。 "
            Call LOG_OUT(LOG_F, "[F110050]終了処理の起動に失敗しました。")
        End If
    End If


    Set F1100101 = Nothing

'Call LOG_OUT(LOG_F, "END 2")

    


    End
End Sub

Private Sub Text2_Change(Index As Integer)
    
    Label3.Caption = ""             '2012.11.06

End Sub

Private Sub Text2_GotFocus(Index As Integer)
    Text2(Index) = Trim(Text2(Index).Text)
    Text2(Index).SelStart = 0
    Text2(Index).SelLength = Len(Text2(Index).Text)
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    
    Text2(0).Text = Format(Val(Text2(0).Text), "00")
    Text2(1).Text = Format(Val(Text2(1).Text), "00")

End Sub

Private Sub Timer1_Timer()
'-------------------------------------------------------
'
'   『終了監視タイマー』
'       2012.05.18
'
'-------------------------------------------------------
    Timer1.Enabled = False
    
    Auto_Off = False                    '2013.06.06
    
    
    If Text2(0).Text & Text2(1).Text = Format(Time, "HHMM") Then
        
        Auto_Off = True                 '2013.06.06
        
        Command1(2).Value = True
    End If
    Timer1.Enabled = True
End Sub

