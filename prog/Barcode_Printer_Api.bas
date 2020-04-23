Attribute VB_Name = "Barcode_Printer_Api"
Option Explicit


'����J�n�֐�(API)
Public Declare Function OpenPrinter& Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) ' Third param changed to long
Public Declare Function StartDocPrinter& Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOC_INFO_1)
Public Declare Function StartPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)

'����ް����������ׂ߰ɑ���֐�(API)
Public Declare Function WritePrinter& Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long)

'����I���֐�(API)
Public Declare Function EndDocPrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
Public Declare Function EndPagePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)
Public Declare Function ClosePrinter& Lib "winspool.drv" (ByVal hPrinter As Long)

'StartDocPrinter�Ŏg�p�����\����
Public Type DOC_INFO_1
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type

'���ʊ֐��װ�萔
Const mDRV_NOERR = 0 '�װ����
Const mDRV_OPENPRINTERERR = 1 'OPENPRINTER�̴װ
Const mDRV_GETPRINTERERR = 2 'GETPRINTER�̴װ
Const mDRV_STATUSCODE_02 = 3 '�������Žð����Ԃ��Ȃ�(�ð���͖���)
Const mDRV_STATUSCODE_04 = 4 'Status3�܂���LanPrinter�ł͂Ȃ�
Const mDRV_STATUSCODE_08 = 5 '�ð���擾���s
Const mDRV_WRITEPRINTER = 6 'WRITEPRINTER�̴װ

'����ް����M�֐�
'����
'plPrinterHandl�c����������
'pvPrintData�c����ް��i�������޲���ް��̍��݂͂ł��܂���j
Public Function PrinterDriver_Write(plPrinterHandl As Long, ByVal pvPrintData As Variant) As Long
    Dim lWritten As Long '�������ײ�ނɑ��M���ꂽ�ް����ނ��ݒ肳���
    Dim bData() As Byte
    Dim sData As String
    
    If "String" = TypeName(pvPrintData) Then
        '�����ް��̏ꍇ
        sData = pvPrintData
        WritePrinter plPrinterHandl, ByVal sData, LenB(StrConv(pvPrintData, vbFromUnicode)), lWritten
        '��ײ�ނɓn���ꂽ�ް����ނ���������
        If lWritten <> LenB(StrConv(pvPrintData, vbFromUnicode)) Then
            '�ُ�I��
            PrinterDriver_Write = mDRV_WRITEPRINTER
            Exit Function
        End If
    Else
        '�޲���ް��̏ꍇ
        bData() = pvPrintData
        WritePrinter plPrinterHandl, bData(0), LenB(pvPrintData), lWritten
        '��ײ�ނɓn���ꂽ�ް����ނ���������
        If lWritten <> LenB(pvPrintData) Then
            '�ُ�I��
            PrinterDriver_Write = mDRV_WRITEPRINTER
            Exit Function
        End If
    End If
    '����I��
    PrinterDriver_Write = mDRV_NOERR
End Function

'����J�n����
'����
'psJobName�c�޷���Ė�
'plPrinterHandl�c���������ق��֐����Ŋ����Ă���
'�߂�l
'0�c����I��
'0�ȊO�c�ُ�I��
Public Function PrinterDriver_Start(ByVal psJobName As String, plPrinterHandl As Long) As Long
Dim lRet    As Long
Dim docinfo As DOC_INFO_1
Dim lJobid  As Long
    
    '���������ق��擾
    lRet = OpenPrinter(Printer.DeviceName, plPrinterHandl, 0)
    If lRet = 0 Then
        '�ُ�I��
        plPrinterHandl = -1
        PrinterDriver_Start = mDRV_OPENPRINTERERR
        Exit Function
    End If
    
    '������J�n����
    docinfo.pDocName = psJobName
    docinfo.pOutputFile = vbNullString
    docinfo.pDatatype = vbNullString
    lJobid = StartDocPrinter(plPrinterHandl, 1, docinfo)
    StartPagePrinter plPrinterHandl
    
    '����I��
    PrinterDriver_Start = mDRV_NOERR
End Function

'����I������
'����
'plPrinterHandl�c����������
Public Sub PrinterDriver_End(plPrinterHandl As Long)
    EndPagePrinter plPrinterHandl
    EndDocPrinter plPrinterHandl
    ClosePrinter plPrinterHandl
End Sub

'���ʊ֐��̴װү���ގ擾�֐�
'�i�hPrinterDriver_�h����n�܂�֐��̖߂�l�ɑΉ�����'�װү���ނ��擾����j
'����
'plNo�c�hPrinterDriver_�h����n�܂�֐��̖߂�l�i�װ�ԍ��j
'�߂�l
'plNo�ɑΉ������װү���ނ��߂����
Public Function PrinterDriver_ErrMsg(plNo As Long) As String
    Dim sMsg As String
    
    Select Case plNo
        Case mDRV_NOERR
            sMsg = ""
        Case mDRV_OPENPRINTERERR
            sMsg = "���������ق̎擾�Ɏ��s���܂����B"
        Case mDRV_GETPRINTERERR
            sMsg = "�ð�����擾�ɍs�����Ƃ��ł��܂���B"
        Case mDRV_STATUSCODE_02
            sMsg = "�������̂��߁A�ð����Ԃ��܂���B"
        Case mDRV_STATUSCODE_04
            sMsg = "Status3 �܂��� LanPrinter�ł͂���܂���B"
        Case mDRV_STATUSCODE_08
            sMsg = "�ð���̎擾�Ɏ��s���܂����B"
        Case mDRV_WRITEPRINTER
            sMsg = "�ް��̑��M�Ɏ��s���܂����B"
        Case Else
            sMsg = "�װ���������܂����B"
    End Select
    
    '�װү���ނ�߂�
    PrinterDriver_ErrMsg = sMsg
End Function

Function Kanji_Conv(psPara As String, psSiftJis As String) As String
Rem ���JIS���ނ���JIS���ނ֕ϊ�
'   psPara      :   H=HEX   B=Binary
'   psSiftJis   :   ���JIS����

Dim i As Integer    '���������ݺ���
Dim vConv           'ܰ��ϐ�
Dim vHex            '4�޲Ă̼��JIS���ނɕϊ������ݺ���
Dim vUpByte         '���2�޲Ă�1�޲Ăɕϊ������ݺ���
Dim vDownByte       '����2�޲Ă�1�޲Ăɕϊ������ݺ���
    
    vConv = ""                                    'ܰ��ϐ��̏�����
    For i = 1 To Len(psSiftJis)                   '�������J��Ԃ�
        vHex = Hex(Asc(Mid$(psSiftJis, i, 1)))    '�S�޲Ă̼��JIS���ނɕϊ�
        If vHex = "20" Then
           Exit For
        End If
        vUpByte = Val("&h" + Mid$(vHex, 1, 2))    '��ʂQ�޲Ă��P�޲Ăɕϊ�
        vDownByte = Val("&h" + Mid$(vHex, 3, 2))  '���ʂQ�޲Ă��P�޲Ăɕϊ�
        If vUpByte >= &HE0 Then                   '��ʂP�޲Ă��d�Oh�̏ꍇ�̏���
           vUpByte = vUpByte - &H40
        End If
        vUpByte = (vUpByte - &H81) * 2 + &H21
        If vDownByte > &H7F Then                  '���ʂP�޲Ă��W�Oh�ȏ�̏���
           vDownByte = vDownByte - 1
        End If
        If vDownByte > &H9D Then                  '���ʂP�޲Ă��X�dh�ȏ�̏���
           vUpByte = vUpByte + 1
           vDownByte = vDownByte - (&H9E - &H21)
        Else
           vDownByte = vDownByte - (&H40 - &H21)  '���ʂP�޲Ă��X�c�ȉ��̏���
        End If
        Select Case psPara
               Case "H"
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ܰ��ϐ��ɑ�������
               Case "B"
                    vConv = vConv + Chr$(vUpByte) + Chr$(vDownByte)  'ܰ��ϐ��ɑ�������
               Case Else
                    vConv = vConv + Hex(vUpByte) + Hex(vDownByte)    'ܰ��ϐ��ɑ�������
        End Select
    Next i
    Kanji_Conv = vConv

End Function

