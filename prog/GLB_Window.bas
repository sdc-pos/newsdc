Attribute VB_Name = "GLB_Window"
Option Explicit

Public hStatusWnd As Long                       '�X�e�[�^�X�o�[�E�B���h�����̃n���h���ۑ��p
'�X�e�[�^�X�o�[�R���g���[�����쐬����
Public Declare Function CreateStatusWindow Lib "comctl32.dll" _
(ByVal style As Long, ByVal lpszText As String, _
    ByVal hwndParent As Long, ByVal wID As Long) As Long
'state�̒萔
Public Const CCS_TOP = &H1                      '�N���C�A���g�̈�̏�ӂɔz�u
Public Const CCS_BOTTOM = &H3                   '���A��ӂɔz�u
Public Const SBARS_SIZEGRIP = &H100             '�T�C�Y�ύX�O���b�v������
'�E�B���h�E�X�^�C���̒萔
Public Const WS_BORDER = &H800000               '�t�H�[���̘g��������
Public Const WS_CAPTION = &HC00000              'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000              '�e�E�C���h�E�����º��۰�(�q�E�C���h�E)���쐬����
Public Const WS_CHILDWINDOW = (WS_CHILD)        '�q�E�C���h�E
Public Const WS_CLIPCHILDREN = &H2000000        '�t�H�[���̍X�V���ɃR���g���[���̍ĕ`���}������
Public Const WS_CLIPSIBLINGS = &H4000000        '�R���g���[���̍X�V���ɑ��̃R���g���[���̍ĕ`���}������
Public Const WS_DISABLED = &H8000000            'Enabled = False
Public Const WS_DLGFRAME = &H400000             '���T�C�Y�ł��Ȃ��g��������
Public Const WS_GROUP = &H20000                 '�R���g���[�� �O���[�v�̍ŏ��̃R���g���[���ł���
Public Const WS_HSCROLL = &H100000              '�����X�N���[���o�[������
Public Const WS_MAXIMIZE = &H1000000            '������Ԃōő剻����
Public Const WS_MAXIMIZEBOX = &H10000           '�ő剻�{�^��������
Public Const WS_MINIMIZE = &H20000000           '������Ԃōŏ�������
Public Const WS_MINIMIZEBOX = &H20000           '�ŏ����{�^��������
Public Const WS_ICONIC = WS_MINIMIZE            'WS_MINIMIZE �Ɠ���
Public Const WS_POPUP = &H80000000              '�|�b�v�A�b�v�^�E�C���h�E
Public Const WS_SYSMENU = &H80000               '�V�X�e�����j���[������
Public Const WS_TABSTOP = &H10000               '�^�u�X�g�b�v�\
Public Const WS_THICKFRAME = &H40000            '���T�C�Y�\�Șg��������
Public Const WS_VISIBLE = &H10000000            'Visible�ł���
Public Const WS_VSCROLL = &H200000              '�����X�N���[���o�[������
Public Const WS_OVERLAPPED = &H0                '�t�H�[���g���ƃL���v�V�����o�[������
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) '�W���I�ȃX�^�C��
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '�|�b�v�A�b�v�^�̕W���I�ȃX�^�C��
Public Const WS_SIZEBOX = WS_THICKFRAME         'WS_THICKFRAME �Ɠ���
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW '32�r�b�g�łł͂�����g�����Ƃ�����
Public Const WS_TILED = WS_OVERLAPPED           'WS_OVERLAPPED �̂Ɠ���
Public Const CW_USEDEFAULT = &H80000000         '�\���ʒu��Windows�����肷��
Public Const WM_USER = &H400                    '���[�U�[����`�ł��郁�b�Z�[�W�̎g�p�̈��\�������ł��ꎩ�̂ɈӖ��͂Ȃ�
'���b�Z�[�W�𑗂�
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageAny Lib "user32.dll" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, lParam As Any) As Long
'=================--�X�e�[�^�X�o�[�֘A===================-
Public Const STATUSCLASSNAME = "msctls_statusbar32"     '�N���X��
'���b�Z�[�W
Public Const SB_SETTEXT = (WM_USER + 1)                 '�e�L�X�g��ݒ肷��
Public Const SB_GETTEXT = (WM_USER + 2)                 '�e�L�X�g���擾����
Public Const SB_GETTEXTLENGTH = (WM_USER + 3)           '�e�L�X�g�̒������擾����
Public Const SB_SETPARTS = (WM_USER + 4)                '�y�C����ݒ肷��
Public Const SB_GETPARTS = (WM_USER + 6)                '�y�C�������擾����
Public Const SB_GETBORDERS = (WM_USER + 7)              '���E���̕����擾����
Public Const SB_SETMINHEIGHT = (WM_USER + 8)            '�E�B���h�E���ŏ��������Ƃ���
                                                        '�X�e�[�^�X�E�B���h�E�̍ŏ��̍���
Public Const SB_SIMPLE = (WM_USER + 9)                  '�V���v���ȃX�^�C���ɂ���
Public Const SB_GETRECT = (WM_USER + 10)                '�w�肳�ꂽ�y�C���̃T�C�Y���擾����
'�e�L�X�g�̕\���̃X�^�C��
Public Const SBT_OWNERDRAW = &H1000                     '�I�[�i�[�`��
Public Const SBT_NOBORDERS = &H100                      '���E���Ȃ�
Public Const SBT_POPOUT = &H200                         '�ʌ`
Public Const SBT_RTLREADING = &H400                     '�E���獶��(�w�u���C��E�A���r�A��̂�)
Public Const SBT_SUNKEN = &H0                           '���`
'���E���̃X�^�C��
Public Const SBB_HORIZONTAL = 0                         '�������E���̕�
Public Const SBB_VERTICAL = 1                           '�������E���̕�
Public Const SBB_DIVIDER = 2                            '�y�C����؂���̕�
'COMCTL32.dll����R�����q�g���[���N���X��o�^����
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Long
Type tagINITCOMMONCONTROLSEX
    dwSize As Long                              '�\���̂̃o�C�g��
    dwICC As Long                               '���[�h����N���X���w�肷��
End Type
Public Const ICC_ANIMATE_CLASS = &H80           '�A�j���[�V�����R���g���[��
Public Const ICC_BAR_CLASSES = &H4              '�c�[���o�[�A�X�e�[�^�X�o�[�A�X���C�_�[�o�[
Public Const ICC_COOL_CLASSES = &H400           '���o�[�R���g���[��
Public Const ICC_DATE_CLASSES = &H100           '�����s�b�N�A�b�v�R���g���[��
Public Const ICC_HOTKEY_CLASS = &H40            '�z�b�g�L�[�R���g���[��
Public Const ICC_INTERNET_CLASSES = &H800       'IP�A�h���X�N���X
Public Const ICC_LISTVIEW_CLASSES = &H1         '���X�g�r���[�A�w�b�_�[�R���g���[��
Public Const ICC_PAGESCROLLER_CLASS = &H1000    '�y�[�W���R���g���[��
Public Const ICC_PROGRESS_CLASS = &H20          '�v���O���X�o�[�R���g���[��
Public Const ICC_TABCLASSES = &H8               '�^�u�R���g���[��
Public Const ICC_TREEVIEW_CLASSES = &H2         '�c���[�r���[�R���g���[��
Public Const ICC_UPDOWN_CLASS = &H10            '�A�b�v�_�E���R���g���[��
Public Const ICC_USEREX_CLASSES = &H200         '�g���R���{�{�b�N�X�N���X
Public Const ICC_WIN95_CLASSES = &HFF           'Windows95�R�����R���g���[��
'�w��̍��W�ʒu�ɂ���E�B���h�E�n���h�����擾����(�X�N���[�����W)
Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As PointApi) As Long
Type PointApi
    X As Long
    Y As Long
End Type
'�E�B���h�E�e�L�X�g�̎擾
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal nMaxCount As String) As Long
'�E�B���h�E�e�L�X�g�T�C�Y�擾
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" _
    (ByVal hwnd As Long) As Long




