Attribute VB_Name = "MainF104019"
Option Explicit

'---------------------------------------------- *更新用在庫ワーク
'ポジショニング
Public wZAIKO_POS   As POSBLK
'データ・バッファ
Public wZAIKOREC    As ZAIKOREC_Tag
'キー・データ
Public K0_wZAIKO    As KEY0_ZAIKO
Public K1_wZAIKO    As KEY1_ZAIKO
Public K2_wZAIKO    As KEY2_ZAIKO





Sub Main()
    Last_JGYOBU = Trim(Command)

    F1040191.Show
End Sub

Public Function wZAIKO_Open(Mode As Integer) As Integer
'****************************************************
'*      「移動処理」    在庫ＯＰＥＮ処理
'*
'*  在庫ファイルを別ポインタでＯＰＥＮする
'*  (呼び元で起動時に１度だけ呼び出す)

'*  戻り値: false       :正常
'*          true        :異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wZAIKO_Open = True
                                '在庫データ　フルパス取込み
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI 読み込みエラー")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- ＯＰＥＮ処理での使用中は、立ち上げ時に１回だけのはずなので、常に画面入力とし、
'               ｷｬﾝｾﾙは、処理の起動ｷｬﾝｾﾙとする。
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    wZAIKO_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "在庫データ")
                Exit Function
        End Select
    Loop

    wZAIKO_Open = False

End Function

Public Function wZAIKO_CLOSE() As Integer

'****************************************************
'*      「移動処理」    在庫ＣＬＯＳＥ処理
'*
'*  在庫ファイルを別ポインタでＣＬＯＳＥする
'*  (呼び元で終了時に１度だけ呼び出す)
'*  戻り値: false       :正常
'*          true        :異常
'****************************************************
Dim sts As Integer
    
    wZAIKO_CLOSE = True
    
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "在庫データ")
            Exit Function
    End Select

    wZAIKO_CLOSE = False

End Function

