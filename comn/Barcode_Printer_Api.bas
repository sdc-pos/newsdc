Attribute VB_Name = "Barcode_Printer_Api"
Option Explicit


'印刷開始関数(API)
Public Declare Function OpenPrinter& Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) ' Third param changed to long
Public Declare Function StartDocPrinter& Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1)
Public Declare Function StartPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)

'印刷ﾃﾞｰﾀをﾌﾟﾘﾝﾀｽﾌﾟｰﾗに送る関数(API)
Public Declare Function WritePrinter& Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long)

'印刷終了関数(API)
Public Declare Function EndDocPrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
Public Declare Function EndPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
Public Declare Function ClosePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)

'StartDocPrinterで使用される構造体
Public Type DOC_INFO_1
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

'共通関数ｴﾗｰ定数
Const mDRV_NOERR = 0 'ｴﾗｰ無し
Const mDRV_OPENPRINTERERR = 1 'OPENPRINTERのｴﾗｰ
Const mDRV_GETPRINTERERR = 2 'GETPRINTERのｴﾗｰ
Const mDRV_STATUSCODE_02 = 3 '処理中でｽﾃｰﾀｽを返せない(ｽﾃｰﾀｽは無効)
Const mDRV_STATUSCODE_04 = 4 'Status3またはLanPrinterではない
Const mDRV_STATUSCODE_08 = 5 'ｽﾃｰﾀｽ取得失敗
Const mDRV_WRITEPRINTER = 6 'WRITEPRINTERのｴﾗｰ

'印刷ﾃﾞｰﾀ送信関数
'引数
'plPrinterHandl…ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙ
'pvPrintData…印刷ﾃﾞｰﾀ（文字とﾊﾞｲﾅﾘﾃﾞｰﾀの混在はできません）
Public Function PrinterDriver_Write(plPrinterHandl As Long, ByVal pvPrintData As Variant) As Long
    Dim lWritten As Long 'ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞに送信されたﾃﾞｰﾀｻｲｽﾞが設定される
    Dim bData() As Byte
    Dim sData As String
    
    If "String" = TypeName(pvPrintData) Then
        '文字ﾃﾞｰﾀの場合
        sData = pvPrintData
        WritePrinter plPrinterHandl, ByVal sData, LenB(StrConv(pvPrintData, vbFromUnicode)), lWritten
        'ﾄﾞﾗｲﾊﾞに渡されたﾃﾞｰﾀｻｲｽﾞをﾁｪｯｸする
        If lWritten <> LenB(StrConv(pvPrintData, vbFromUnicode)) Then
            '異常終了
            PrinterDriver_Write = mDRV_WRITEPRINTER
            Exit Function
        End If
    Else
        'ﾊﾞｲﾅﾘﾃﾞｰﾀの場合
        bData() = pvPrintData
        WritePrinter plPrinterHandl, bData(0), LenB(pvPrintData), lWritten
        'ﾄﾞﾗｲﾊﾞに渡されたﾃﾞｰﾀｻｲｽﾞをﾁｪｯｸする
        If lWritten <> LenB(pvPrintData) Then
            '異常終了
            PrinterDriver_Write = mDRV_WRITEPRINTER
            Exit Function
        End If
    End If
    '正常終了
    PrinterDriver_Write = mDRV_NOERR
End Function

'印刷開始処理
'引数
'psJobName…ﾄﾞｷｭﾒﾝﾄ名
'plPrinterHandl…ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙが関数内で割当てられる
'戻り値
'0…正常終了
'0以外…異常終了
Public Function PrinterDriver_Start(ByVal psJobName As String, plPrinterHandl As Long) As Long
Dim lRet    As Long
Dim docinfo As DOC_INFO_1
Dim lJobid  As Long
    
    'ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙを取得
    lRet = OpenPrinter(Printer.DeviceName, plPrinterHandl, 0)
    If lRet = 0 Then
        '異常終了
        plPrinterHandl = -1
        PrinterDriver_Start = mDRV_OPENPRINTERERR
        Exit Function
    End If
    
    '印刷を開始する
    docinfo.pDocName = psJobName
    docinfo.pOutputFile = vbNullString
    docinfo.pDatatype = vbNullString
    lJobid = StartDocPrinter(plPrinterHandl, 1, docinfo)
    StartPagePrinter plPrinterHandl
    
    '正常終了
    PrinterDriver_Start = mDRV_NOERR
End Function

'印刷終了処理
'引数
'plPrinterHandl…ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙ
Public Sub PrinterDriver_End(plPrinterHandl As Long)
    EndPagePrinter plPrinterHandl
    EndDocPrinter plPrinterHandl
    ClosePrinter plPrinterHandl
End Sub

'共通関数のｴﾗｰﾒｯｾｰｼﾞ取得関数
'（”PrinterDriver_”から始まる関数の戻り値に対応した'ｴﾗｰﾒｯｾｰｼﾞを取得する）
'引数
'plNo…”PrinterDriver_”から始まる関数の戻り値（ｴﾗｰ番号）
'戻り値
'plNoに対応したｴﾗｰﾒｯｾｰｼﾞが戻される
Public Function PrinterDriver_ErrMsg(plNo As Long) As String
    Dim sMsg As String
    
    Select Case plNo
        Case mDRV_NOERR
            sMsg = ""
        Case mDRV_OPENPRINTERERR
            sMsg = "ﾌﾟﾘﾝﾀﾊﾝﾄﾞﾙの取得に失敗しました。"
        Case mDRV_GETPRINTERERR
            sMsg = "ｽﾃｰﾀｽを取得に行くことができません。"
        Case mDRV_STATUSCODE_02
            sMsg = "処理中のため、ｽﾃｰﾀｽを返せません。"
        Case mDRV_STATUSCODE_04
            sMsg = "Status3 または LanPrinterではありません。"
        Case mDRV_STATUSCODE_08
            sMsg = "ｽﾃｰﾀｽの取得に失敗しました。"
        Case mDRV_WRITEPRINTER
            sMsg = "ﾃﾞｰﾀの送信に失敗しました。"
        Case Else
            sMsg = "ｴﾗｰが発生しました。"
    End Select
    
    'ｴﾗｰﾒｯｾｰｼﾞを戻す
    PrinterDriver_ErrMsg = sMsg
End Function

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

