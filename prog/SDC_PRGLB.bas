Attribute VB_Name = "SDC_PRGLB"
Option Explicit
'�t�H���g��
Global Const SDC_PGL_FONT_GOTHIC = "�l�r �S�V�b�N"
Global Const SDC_PGL_FONT_PGOTHIC = "�l�r �o�S�V�b�N"
Global Const SDC_PGL_FONT_AGOTHIC = "@�l�r �S�V�b�N"
Global Const SDC_PGL_FONT_MINCYO = "�l�r ����"
Global Const SDC_PGL_FONT_PMINCYO = "�l�r �o����"
Global Const SDC_PGL_FONT_AMINCYO = "@�l�r ����"
Global Const SDC_PGL_FONT_BCODE39 = "3 of 9 Barcode"

'����p���`��
'vbPRPSA3               A3�A297 x 420 mm
Global Const SDC_PGL_SHEET_A3V% = 0          '�c
Global Const SDC_PGL_SHEET_A3H% = 1          '��
'vbPRPSA4               A4�A210 x 297 mm
Global Const SDC_PGL_SHEET_A4V% = 2          '�c
Global Const SDC_PGL_SHEET_A4H% = 3          '��
'vbPRPSA4Small          A4 Small�A210 x 297 mm
Global Const SDC_PGL_SHEET_A4SV% = 4         '�c
Global Const SDC_PGL_SHEET_A4SH% = 5         '��
'vbPRPSA5               A5�A148 x 210 mm
Global Const SDC_PGL_SHEET_A5V% = 6          '�c
Global Const SDC_PGL_SHEET_A5H% = 7          '��
'vbPRPSB4               B4�A250 x 354 mm
Global Const SDC_PGL_SHEET_B4V% = 8          '�c
Global Const SDC_PGL_SHEET_B4H% = 9          '��
'vbPRPSB5               B5�A182 x 257 mm
Global Const SDC_PGL_SHEET_B5V% = 10         '�c
Global Const SDC_PGL_SHEET_B5H% = 11         '��
'vbPRPSLetter           ���^�[�A8 1/2 x 11 �C���`
Global Const SDC_PGL_SHEET_LETV% = 12        '�c
Global Const SDC_PGL_SHEET_LETH% = 13        '��
'vbPRPSLetterSmall      ���^�[ �X���[���A8 1/2 x 11 �C���`
Global Const SDC_PGL_SHEET_LTSV% = 14        '�c
Global Const SDC_PGL_SHEET_LTSH% = 15        '��
'vbPRPSTabloid          �^�u���C�h�A11 x 17 �C���`
Global Const SDC_PGL_SHEET_TABV% = 16        '�c
Global Const SDC_PGL_SHEET_TABH% = 17        '��
'vbPRPSLedger           ���W���[�A17 x 11 �C���`
Global Const SDC_PGL_SHEET_LEDV% = 18        '�c
Global Const SDC_PGL_SHEET_LEDH% = 19        '��
'vbPRPSLegal            ���[�K���A8 1/2 x 14 �C���`
Global Const SDC_PGL_SHEET_LEGV% = 20        '�c
Global Const SDC_PGL_SHEET_LEGH% = 21        '��
'vbPRPSStatement        �X�e�[�g�����g�A5 1/2 x 8 1/2 �C���`
Global Const SDC_PGL_SHEET_STMV% = 22        '�c
Global Const SDC_PGL_SHEET_STMH% = 23        '��
'vbPRPSExecutive        �G�O�[�N�e�B�u�A7 1/2 x 10 1/2 �C���`
Global Const SDC_PGL_SHEET_EXEV% = 24        '�c
Global Const SDC_PGL_SHEET_EXEH% = 25        '��
'vbPRPSFolio            �t�H���I�A8 1/2 x 13 �C���`
Global Const SDC_PGL_SHEET_FOLV% = 26        '�c
Global Const SDC_PGL_SHEET_FOLH% = 27        '��
'vbPRPSQuarto           �N�H�[�g�A215 x 275 mm
Global Const SDC_PGL_SHEET_QUAV% = 28        '�c
Global Const SDC_PGL_SHEET_QUAH% = 29        '��
'vbPRPS10x14            10 x 14 �C���`
Global Const SDC_PGL_SHEET_10x14V% = 30      '�c
Global Const SDC_PGL_SHEET_10x14H% = 31      '��
'vbPRPS11x17            11 x 17 �C���`
Global Const SDC_PGL_SHEET_11x17V% = 32      '�c
Global Const SDC_PGL_SHEET_11x17H% = 33      '��
'vbPRPSNote             �m�[�g�A8 1/2 x 11 �C���`
Global Const SDC_PGL_SHEET_NOTV% = 34        '�c
Global Const SDC_PGL_SHEET_NOTH% = 35        '��
'vbPRPSCSheet           C �T�C�Y �V�[�g
Global Const SDC_PGL_SHEET_CV% = 36          '�c
Global Const SDC_PGL_SHEET_CH% = 37          '��
'vbPRPSDSheet           D �T�C�Y �V�[�g
Global Const SDC_PGL_SHEET_DV% = 38          '�c
Global Const SDC_PGL_SHEET_DH% = 39          '��
'vbPRPSESheet           E �T�C�Y �V�[�g
Global Const SDC_PGL_SHEET_EV% = 40          '�c
Global Const SDC_PGL_SHEET_EH% = 41          '��
'vbPRPSFanfoldUS        U.S. ����ް�� ̧�̫���ށA14 7/8 x 11 ���
Global Const SDC_PGL_SHEET_USV% = 42         '�c
Global Const SDC_PGL_SHEET_USH% = 43         '��
'vbPRPSUser             ���[�U�[��`
Global Const SDC_PGL_SHEET_USRV% = 44        '�c
Global Const SDC_PGL_SHEET_USRH% = 45        '��


'����p���[�N
Global Const SDC_PGL_LINI% = 99              '�s���J�E���^�����l
Global SDC_PGL_Lcnt As Integer               '�s���J�E���^

Global SDC_PGL_Pdate As String               '����J�n���t�iͯ�ް�p�j
Global SDC_PGL_Ptime As String               '����J�n�����iͯ�ް�p�j

Global SDC_PGL_PRT_CAN As Boolean            '�����ݾ� �׸�

Function SDC_PGL_Init(Printr_Nm As String, Font_Nm As String, Font_Siz As Integer, Sheet_Type As Integer) As Integer
'----------------------------------------------------------------------
'�@�@�@�v�����^�[�@�����ݒ�
'
'  Printr_Nm �F�v�����^���擾�p�L�[������
'  Font_Nm   �F�t�H���g���iNull�l�Ȃ�v�����^���̐ݒ�̂݁j
'  Font_Siz  �F�t�H���g�T�C�Y
'  Sheet_Type�F����p���`�ԁi�ڍׂ͖{Ӽޭ�ق�Global��`�Q�Ɓj
'
'�@�߂�l�F�Ȃ�
'          CREATE 1999.04.17  S.Shibano
'----------------------------------------------------------------------
Dim Wk_Printer As PRINTER
Dim sts As Integer
Dim c As String
Dim USE_PRINTER As String

    SDC_PGL_Init = True

'�w�蒠�[�p�v�����^���@�擾
    If GetIni("PRINTER", "SYSTEM", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���v�����^����`����Ă��܂���B", vbCritical
        Exit Function
    End If
    USE_PRINTER = RTrim(c)      '��̫�ľ��

    If GetIni("PRINTER", Printr_Nm, "SYS", c) = False Then
        USE_PRINTER = RTrim(c)
    Else
        Beep
        MsgBox Printr_Nm & "�p�v�����^�̐ݒ�l(SYS.INI)����", vbExclamation
        Exit Function
    End If

'�w�蒠�[�p�v�����^���擾
    For Each Wk_Printer In Printers
        c = RTrim(Wk_Printer.DeviceName)
        If c = USE_PRINTER Then
            Set PRINTER = Wk_Printer
            Exit For
        End If
    Next
    If Font_Nm = "" Then
        SDC_PGL_Init = False
        Exit Function
    End If
'����t�H���g�ݒ�
    Call SDC_PGL_Font(Font_Nm, Font_Siz)

'����p���`�ԁ@�ݒ�
    Select Case Sheet_Type

        Case SDC_PGL_SHEET_A3V             '�`�R�c
            PRINTER.PaperSize = vbPRPSA3
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A3H             '�`�R��
            PRINTER.PaperSize = vbPRPSA3
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A4V             '�`�S�c
            PRINTER.PaperSize = vbPRPSA4
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A4H             '�`�S��
            PRINTER.PaperSize = vbPRPSA4
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A4SV            'A4 Small�c
            PRINTER.PaperSize = vbPRPSA4Small
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A4SH            'A4 Small��
            PRINTER.PaperSize = vbPRPSA4Small
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A5V             '�`�T�c
            PRINTER.PaperSize = vbPRPSA5
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_A5H             '�`�T��
            PRINTER.PaperSize = vbPRPSA5
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_B4V             '�a�S�c
            PRINTER.PaperSize = vbPRPSB4
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_B4H             '�a�S��
            PRINTER.PaperSize = vbPRPSB4
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_B5V             '�a�T�c
            PRINTER.PaperSize = vbPRPSB5
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_B5H             '�a�T��
            PRINTER.PaperSize = vbPRPSB5
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LETV            '���^�[�c
            PRINTER.PaperSize = vbPRPSLetter
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LETH            '���^�[��
            PRINTER.PaperSize = vbPRPSLetter
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LTSV            '���^�[ �X���[���c
            PRINTER.PaperSize = vbPRPSLetterSmall
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LTSH            '���^�[ �X���[����
            PRINTER.PaperSize = vbPRPSLetterSmall
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_TABV             '�^�u���C�h�c
            PRINTER.PaperSize = vbPRPSTabloid
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_TABH             '�^�u���C�h��
            PRINTER.PaperSize = vbPRPSTabloid
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LEDV             '���W���[�c
            PRINTER.PaperSize = vbPRPSLedger
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LEDH             '���W���[��
            PRINTER.PaperSize = vbPRPSLedger
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LEGV             '���[�K���c
            PRINTER.PaperSize = vbPRPSLegal
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_LEGH             '���[�K����
            PRINTER.PaperSize = vbPRPSLegal
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_STMV             '�X�e�[�g�����g�c
            PRINTER.PaperSize = vbPRPSStatement
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_STMH             '�X�e�[�g�����g��
            PRINTER.PaperSize = vbPRPSStatement
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_EXEV             '�G�O�[�N�e�B�u�c
            PRINTER.PaperSize = vbPRPSExecutive
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_EXEH             '�G�O�[�N�e�B�u��
            PRINTER.PaperSize = vbPRPSExecutive
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_FOLV             '�t�H���I�c
            PRINTER.PaperSize = vbPRPSFolio
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_FOLH             '�t�H���I�c
            PRINTER.PaperSize = vbPRPSFolio
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_QUAV             '�N�H�[�g�c
            PRINTER.PaperSize = vbPRPSQuarto
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_QUAH             '�N�H�[�g��
            PRINTER.PaperSize = vbPRPSQuarto
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_10x14V             '10 x 14�c
            PRINTER.PaperSize = vbPRPS10x14
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_10x14H             '10 x 14��
            PRINTER.PaperSize = vbPRPS10x14
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_11x17V             '11 x 17�c
            PRINTER.PaperSize = vbPRPS11x17
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_11x17H             '11 x 17��
            PRINTER.PaperSize = vbPRPS11x17
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_NOTV             '�m�[�g�c
            PRINTER.PaperSize = vbPRPSNote
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_NOTH             '�m�[�g��
            PRINTER.PaperSize = vbPRPSNote
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_CV             'C �T�C�Y�c
            PRINTER.PaperSize = vbPRPSCSheet
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_CH             'C �T�C�Y��
            PRINTER.PaperSize = vbPRPSCSheet
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_DV             'D �T�C�Y�c
            PRINTER.PaperSize = vbPRPSDSheet
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_DH             'D �T�C�Y��
            PRINTER.PaperSize = vbPRPSDSheet
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_EV             'E �T�C�Y�c
            PRINTER.PaperSize = vbPRPSESheet
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_EH             'E �T�C�Y��
            PRINTER.PaperSize = vbPRPSESheet
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_USV             'U.S. ����ް�ޏc
            PRINTER.PaperSize = vbPRPSFanfoldUS
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��

        Case SDC_PGL_SHEET_USH             'U.S. ����ް�މ�
            PRINTER.PaperSize = vbPRPSFanfoldUS
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��
            
        Case SDC_PGL_SHEET_USRV              '���[�U��`�@�c
            'Printer.PaperSize = vbPRPSUser
            PRINTER.Orientation = vbPRORPortrait   '�p���̒Z�ӂ���ɂ��Ĉ��
            
        Case SDC_PGL_SHEET_USRH              '���[�U��`�@��
            'Printer.PaperSize = vbPRPSUser
            PRINTER.Orientation = vbPRORLandscape  '�p���̒��ӂ���ɂ��Ĉ��
            
        Case Else
        
    End Select

'����p���[�N�����ݒ�
    SDC_PGL_Lcnt = SDC_PGL_LINI         '�s���J�E���^������
    SDC_PGL_Pdate = Date                '����J�n�@���t
    SDC_PGL_Ptime = Time                '�@�@�@�@�@����

    SDC_PGL_Init = False

End Function

Sub SDC_PGL_Font(Font_Nm As String, Font_Siz As Integer)
'----------------------------------------------------------------------
'�@�@�@             �t�H���g�ݒ�
'
'  Font_Nm   �F�t�H���g��
'  Font_Siz  �F�t�H���g�T�C�Y
'
'          CREATE 1999.05.28  S.Shibano
'----------------------------------------------------------------------
Dim W_Font As New StdFont

'����t�H���g�ݒ�
    With W_Font
        .NAME = Font_Nm
        .Size = Font_Siz
    End With
    Set PRINTER.Font = W_Font

End Sub
