Attribute VB_Name = "F900070bas"
Option Explicit

Public OutFno   As Integer
Public DataNo   As Integer

Public ZaikoData   As Variant


'レコード定義
Private Type OutREC_Tag
    TEXT_NO(0 To 8) As Byte         'ﾃｷｽﾄ№
    JGYOBU(0 To 0) As Byte          '事業部区分
    CYOK_KBN(0 To 0) As Byte        '直送区分
    DEN_DT(0 To 7) As Byte          '伝票日付
    IO_KBN(0 To 0) As Byte          '入出庫区分
    PM_KBN(0 To 0) As Byte          '赤黒区分
    DEN_SYU(0 To 0) As Byte         '伝票種別
    DEN_NO(0 To 5) As Byte          '伝票№
    CYU_KBN(0 To 0) As Byte         '注文区分
    HIN_GAI(0 To 12) As Byte        '品番（外部）
    HIN_NAI(0 To 12) As Byte        '品番（内部）
    HIN_NAME(0 To 24) As Byte       '品名
    YOTEI_QTY(0 To 5) As Byte       '数量
    YOSAN_FROM(0 To 4) As Byte      '予算単位（元）
    YOSAN_TO(0 To 4) As Byte        '予算単位（先）
    HOST_SOKO(0 To 1) As Byte       '倉庫区分（ﾎｽﾄ）
    HOST_TANA(0 To 7) As Byte       '棚番（ﾎｽﾄ）
    SYUK_CODE(0 To 4) As Byte       '支給先／出荷先
    SYUK_NAME(0 To 19) As Byte      '支給先／出荷先名
    REC_END(0 To 0) As Byte         'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    CR_LF(0 To 1) As Byte           'CR.LF
End Type

'データ・バッファ
Public OutREC As OutREC_Tag

Function OutREC_Open_Proc() As Integer
Dim ans As Integer
Dim sts As Integer

    
    On Error Resume Next    'ｴﾗｰﾄﾗｯﾌﾟON
    
    
    Kill "c:\zaiko\shiji79.dat"
    OutREC_Open_Proc = False
    OutFno = FreeFile
    
    On Error GoTo OutREC_Open_Err    'ｴﾗｰﾄﾗｯﾌﾟON

    Open "c:\zaiko\shiji79.dat" For Binary As #OutFno

    Exit Function

OutREC_Open_Err:     'ｴﾗｰ処理ﾙｰﾁﾝ
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Beep
    ans = MsgBox("エラー [在庫移管データ : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    OutREC_Open_Proc = True
End Function

Public Function Data_Put_Proc() As Integer
    
Dim wk13    As String * 13
Dim wk25    As String * 25
    
    Data_Put_Proc = True
    
    
    On Error GoTo Error_Proc
    
    DataNo = DataNo + 1
                                'ﾃｷｽﾄ№
    Call UniCode_Conv(OutREC.TEXT_NO, (Format(DataNo, "000000000")))
                                '事業部
    Call UniCode_Conv(OutREC.JGYOBU, "7")
                                '直送区分
    Call UniCode_Conv(OutREC.CYOK_KBN, "0")
                                '伝票日付
    Call UniCode_Conv(OutREC.DEN_DT, Format(Now, "YYYYMMDD"))
                                '入出庫区分
    Call UniCode_Conv(OutREC.IO_KBN, "1")
                                '赤黒区分
    Call UniCode_Conv(OutREC.PM_KBN, "")
                                '伝票種別
    Call UniCode_Conv(OutREC.DEN_SYU, "")
                                '伝票№
    Call UniCode_Conv(OutREC.DEN_NO, Format(DataNo, "000000"))
                                '注文区分
    Call UniCode_Conv(OutREC.CYU_KBN, " ")
                                '品番外部
    
    wk13 = ZaikoData(0)
    Call UniCode_Conv(OutREC.HIN_GAI, wk13)
                                '品番内部
    wk13 = ZaikoData(2)
    Call UniCode_Conv(OutREC.HIN_NAI, wk13)
                                '品名
    wk25 = ZaikoData(1)
    Call UniCode_Conv(OutREC.HIN_NAME, wk25)
                                '数量
    If Not IsNumeric(ZaikoData(3)) Then
        Call UniCode_Conv(OutREC.YOTEI_QTY, "000000")
    Else
        Call UniCode_Conv(OutREC.YOTEI_QTY, Format(CInt(ZaikoData(3)), "000000"))
    End If
                                '予算単位（元）
    Call UniCode_Conv(OutREC.YOSAN_FROM, "")
                                '予算単位（先）
    Call UniCode_Conv(OutREC.YOSAN_TO, "")
                                '倉庫区分（ホスト）
    Call UniCode_Conv(OutREC.HOST_SOKO, "K ")
                                '棚番（ホスト）
    Call UniCode_Conv(OutREC.HOST_TANA, "")
                                '支給先／出荷先
    Call UniCode_Conv(OutREC.SYUK_CODE, "")
                                '支給先／出荷先名
    Call UniCode_Conv(OutREC.SYUK_NAME, "")
                                'レコード終端マーク（@）
    Call UniCode_Conv(OutREC.REC_END, "@")
    Call UniCode_Conv(OutREC.CR_LF, vbCrLf)
    
    Put #OutFno, , OutREC

    Data_Put_Proc = False
    Exit Function

Error_Proc:

End Function
