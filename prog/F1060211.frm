VERSION 5.00
Begin VB.Form F1060211 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�u���i�����ёΉ��v���i���v��x���A���[�����X�g��� "
   ClientHeight    =   6948
   ClientLeft      =   2328
   ClientTop       =   2712
   ClientWidth     =   11292
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6948
   ScaleWidth      =   11292
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   7
      Left            =   8715
      MaxLength       =   2
      TabIndex        =   34
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   8085
      MaxLength       =   2
      TabIndex        =   32
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   5
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   30
      Top             =   3600
      Width           =   645
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   4
      Left            =   6300
      MaxLength       =   2
      TabIndex        =   28
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   3
      Left            =   5670
      MaxLength       =   2
      TabIndex        =   26
      Top             =   3600
      Width           =   330
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   2
      Left            =   4620
      MaxLength       =   4
      TabIndex        =   24
      Top             =   3600
      Width           =   645
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   1
      Left            =   4620
      TabIndex        =   21
      Top             =   2760
      Width           =   330
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   4620
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   19
      Top             =   2160
      Width           =   1170
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   5460
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   17
      Top             =   1440
      Width           =   3270
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   4620
      TabIndex        =   16
      Top             =   1440
      Width           =   750
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4575
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   840
      Width           =   1125
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   11
      Left            =   9030
      TabIndex        =   35
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   10
      Left            =   8400
      TabIndex        =   33
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   9
      Left            =   7770
      TabIndex        =   31
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���`"
      Height          =   255
      Index           =   8
      Left            =   6615
      TabIndex        =   29
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   7
      Left            =   6090
      TabIndex        =   27
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   6
      Left            =   5355
      TabIndex        =   25
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ώ۔N����"
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   23
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i���󔒁@�S�q�Ɏw��j"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   22
      Top             =   2880
      Width           =   2730
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�W���I�ԁi�q�ɔԍ��j"
      Height          =   255
      Index           =   3
      Left            =   1995
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����i�݌�"
      Height          =   255
      Index           =   2
      Left            =   3150
      TabIndex        =   18
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�󕥐�R�[�h"
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����敪"
      Height          =   255
      Index           =   0
      Left            =   3210
      TabIndex        =   14
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1060211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxUKEHARAI_CODE% = 0         '�󕥐�R�[�h
Private Const ptxSOKO_NO% = 1               '�q�ɔԍ�
Private Const ptxS_YY% = 2                  '�J�n�N�����@�N
Private Const ptxS_MM% = 3                  '�J�n�N�����@��
Private Const ptxS_DD% = 4                  '�J�n�N�����@��
Private Const ptxE_YY% = 5                  '�I���N�����@�N
Private Const ptxE_MM% = 6                  '�I���N�����@��
Private Const ptxE_DD% = 7                  '�I���N�����@��


Private Const Text_Max% = 7                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbTORI_KBN% = 0             '�����R�[�h
Private Const pcmbUKEHARAI_CODE% = 1        '�󕥐�R�[�h
Private Const pcmbMI_ZAIKO% = 2             '�����i�݌�


Private Const LMAX% = 36                    '�œ��ő�s��
Private Const LCTL% = 99                    '
Private Const MGN_L% = 10                   '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j


Private NormalFont  As New StdFont          '����t�H���g
Private MidFont     As New StdFont          '����t�H���g

Private OutSide     As Long                 '����ΊO�o�א�

Private GOODS_DATA  As String               '�o�̓f�[�^�t�@�C����

Private NON_MTS     As String               '���O������


Private Type EE_ZAIKO_TBL_tag
    EE_LOC          As String * 8
    EE_QTY          As Long
End Type

Private EE_ZAIKO_TBL(0 To 2) As EE_ZAIKO_TBL_tag

Private SHO_SOKO    As Variant              '���i���p�q��(�����i�Ƃ݂Ȃ���)

Private Const Last_Update_day$ = "([F106021] 2011.07.14 12:00)"

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   �G���[�`�F�b�N����
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
            
            
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060211.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060211)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060211)


    F1060211.MousePointer = vbDefault

End Sub



Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)


    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
    
        Case pcmbTORI_KBN
        
        
        
            If Ukeharai_Set_Proc() Then
                Unload Me
            End If
        
            Combo(pcmbUKEHARAI_CODE).SetFocus
        
        
        
        Case pcmbUKEHARAI_CODE
        
            Text(ptxUKEHARAI_CODE).Text = Right(Combo(pcmbUKEHARAI_CODE).Text, 5)
            Combo(pcmbMI_ZAIKO).SetFocus
        
        
        Case pcmbMI_ZAIKO
    
            Text(ptxSOKO_NO).SetFocus
    
    End Select


End Sub

Private Sub Combo_LostFocus(Index As Integer)

Dim i   As Integer


    Select Case Index


        Case pcmbTORI_KBN



            If Ukeharai_Set_Proc() Then
                Unload Me
            End If

'            Combo(pcmbUKEHARAI_CODE).SetFocus

            If Trim(Text(ptxUKEHARAI_CODE).Text) <> "" Then
                For i = 0 To Combo(pcmbUKEHARAI_CODE).ListCount - 1
                    If Trim(Text(ptxUKEHARAI_CODE).Text) = Trim(Right(Combo(pcmbUKEHARAI_CODE).List(i), 5)) Then
                        Combo(pcmbUKEHARAI_CODE).ListIndex = i
                        Exit For
                    End If
                Next i
            End If




        Case pcmbUKEHARAI_CODE

            Text(ptxUKEHARAI_CODE).Text = Right(Combo(pcmbUKEHARAI_CODE).Text, 5)
'            Combo(pcmbMI_ZAIKO).SetFocus


        Case pcmbMI_ZAIKO

'            Text(ptxSOKO_NO).SetFocus
    
    End Select



End Sub

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
    Select Case Index
        
        Case 7                              '�f�[�^�o��
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                If Data_Proc() Then
                    Unload Me
                End If
            End If
            
            Text(ptxS_YY).SetFocus
        
        
        Case 8                              '���
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
            
            Beep
            yn = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxS_YY).SetFocus
                    
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Form_DblClick()
     PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

'
Private Sub Form_Load()

Dim c   As String * 128
Dim i   As Integer
     
     If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1060211.Caption = "���i���v��x���A���[�����X�g(�o���ް��Ή�)����i" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                '���i���x���t�@�C������荞��
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'���i���x���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GOODS_DATA = Trim(c)
                                
'-----------    SYS.INI ---> (����۸���ID).INI 2011.07.14
                                '�ΏۊO�o�א���荞��
    If GetIni(App.EXEName, "OUTSIDE", App.EXEName, c) Then
        OutSide = 0
    Else
        If IsNumeric(Trim(c)) Then
            OutSide = CLng(Trim(c))
        Else
            OutSide = 0
        End If
    End If
                                '���i���p�q�Ɏ�荞��
    If GetIni(App.EXEName, "SHO_SOKO", "SYS", c) Then
        c = " "
    End If
    SHO_SOKO = Split(Trim(c), ",", -1)
                                
'-----------    SYS.INI ---> (����۸���ID).INI 2011.07.14
                                
                                
                                '���O�������荞��
    If GetIni("PI00010", "MTSSS", "P_SYS", c) Then
        NON_MTS = ""
    Else
        NON_MTS = Trim(c)
    End If
                                
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���W�v�t�@�C���n�o�d�m
    If GOODS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���w�}�f�[�^�i�e�j�n�o�d�m 2007.11.14
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�(�ʏ�)
    With NormalFont
        .NAME = F1060211.FontName
        .Size = 12
    End With

                                '����t�H���g�ݒ�i���j
    With MidFont
        .NAME = F1060211.FontName
        .Size = 8
    End With





    Combo(pcmbTORI_KBN).AddItem "�S�@�ā@     " & " "
    Combo(pcmbTORI_KBN).AddItem P_TORI_GENERAL_N & "     " & P_TORI_GENERAL
    Combo(pcmbTORI_KBN).AddItem P_TORI_NAISYOKU_N & "     " & P_TORI_NAISYOKU
    Combo(pcmbTORI_KBN).AddItem P_TORI_GENKIN_N & "     " & P_TORI_GENKIN
    Combo(pcmbTORI_KBN).AddItem P_TORI_SYANAI_N & "     " & P_TORI_SYANAI
    Combo(pcmbTORI_KBN).AddItem P_TORI_ANOTHER_N & "     " & P_TORI_ANOTHER
    Combo(pcmbTORI_KBN).AddItem P_TORI_JIKYU_N & "     " & P_TORI_JIKYU
    Combo(pcmbTORI_KBN).ListIndex = 0

    Combo(pcmbMI_ZAIKO).AddItem "�S�@�ā@     " & "0"
    Combo(pcmbMI_ZAIKO).AddItem "�O�����@     " & "1"
    Combo(pcmbMI_ZAIKO).AddItem "�O�̂݁@     " & "2"
    Combo(pcmbMI_ZAIKO).ListIndex = 0


    

    Show
    
    Combo(pcmbTORI_KBN).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
                                            '���i���W�v�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
                                            '���i���w�}�f�[�^(�e)�b�k�n�r�d
    sts = BTRV(BtOpClose, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���w�}�f�[�^(�e)")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060211 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1060211.Caption = "���i���v��x���A���[�����X�g(�o���ް��Ή�)����i" + RTrim(JGYOBU_T(Index).NAME) + "�j" & Last_Update_day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
Dim sts As Integer

    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Select Case Index
        
        Case ptxUKEHARAI_CODE
        
            For i = 0 To Combo(pcmbUKEHARAI_CODE).ListCount - 1
                If Trim(Text(ptxUKEHARAI_CODE).Text) = Trim(Right(Combo(pcmbUKEHARAI_CODE).List(i), 5)) Then
                    Combo(pcmbUKEHARAI_CODE).ListIndex = i
                    Exit For
                End If
            Next i
        
            If i > Combo(pcmbUKEHARAI_CODE).ListCount - 1 Then
        
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text(Index).SetFocus
                Exit Sub
        
            End If
        
        Case ptxSOKO_NO
        
        Case ptxS_YY
        
                    
        Case ptxS_MM, ptxS_DD
    
            If Trim(Text(Index).Text) = "" Then
            Else
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
        Case ptxE_YY
        
                    
        Case ptxE_MM, ptxE_DD
    
            
            
            If Trim(Text(Index).Text) = "" Then
            Else
            
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
        
            If Index = ptxE_DD Then
                If Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text > _
                    Text(ptxE_YY).Text & Text(ptxE_MM).Text & Text(ptxE_DD).Text Then
    
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text(ptxS_YY).SetFocus
                    Exit Sub
                End If
            End If
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���i���x���A���[�����X�g�������
'----------------------------------------------------------------------------
Dim LCNT        As Integer

Dim sts         As Integer
Dim com         As Integer
Dim FSW         As Boolean

Dim Save_Soko   As String * 2

Dim Edit        As String

Dim SKIP_Flg    As Boolean

Dim X_Tab       As Integer

    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If



    LCNT = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    Call UniCode_Conv(K1_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K1_GOODS.ST_RETU, "")
    Call UniCode_Conv(K1_GOODS.ST_REN, "")
    Call UniCode_Conv(K1_GOODS.ST_DAN, "")
    Call UniCode_Conv(K1_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K1_GOODS.HIN_GAI, "")
    
    
    com = BtOpGetGreater
    FSW = True
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select


'-------------------------------------------------  ���׈��
            
        SKIP_Flg = False
        Select Case Right(Combo(pcmbMI_ZAIKO).Text, 1)
            Case "0"        '�S��
            Case "1"        '0�ΏۊO
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) = 0 Then
                    SKIP_Flg = True
                End If
            Case "2"        '0�̂�
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <> 0 Then
                    SKIP_Flg = True
                End If
        End Select
            
            
            
            
        
        If SKIP_Flg Then
        Else
            
            If FSW Then
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        End If
                    Case BtErrKeyNotFound
                        '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                FSW = False
            End If
            
            
            
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                LCNT = LMAX + 1
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        End If
                    
                    Case BtErrKeyNotFound
                            '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                
            End If
            
            
            
            
            
            If Head_Print_Proc(LCNT) Then
                Exit Function
            End If
        
            X_Tab = MGN_L
        
            Printer.Print Tab(X_Tab);
                                                    '�W���I��
            Edit = StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
            Printer.Print Edit;
            X_Tab = X_Tab + Len(Edit) + 5
'                X_Tab = X_Tab + Len(Edit) + 3
                                                    '�i�ԁi�O���j
            Printer.Print Tab(X_Tab);

            Printer.Print Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 5
            X_Tab = X_Tab + Len(Left(StrConv(GOODSREC.HIN_GAI, vbUnicode), 13)) + 4
                                                    '����
            Printer.Print Tab(X_Tab);
            Printer.Print StrConv(GOODSREC.PACKING_NO, vbUnicode);
'                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 5
            X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                                                    '���i���ςݍ݌ɐ�
            Printer.Print Tab(X_Tab);
            Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
            X_Tab = X_Tab + Len(Edit) + 2
                                                    '�����i�݌ɐ�
            Printer.Print Tab(X_Tab);
            Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
            X_Tab = X_Tab + Len(Edit) + 2
                                                    '�����Ϗo�א�
            Printer.Print Tab(X_Tab);
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
            X_Tab = X_Tab + Len(Edit) + 2
                                                    '���O���i���K�v��
            Printer.Print Tab(X_Tab);
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Printer.Print Edit;
'                X_Tab = X_Tab + Len(Edit) + 8
            X_Tab = X_Tab + Len(Edit) + 2
                                                    '���O���i����
            Printer.Print Tab(X_Tab);
            Edit = Format(CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If

            Printer.Print Edit;
            X_Tab = X_Tab + Len(Edit) + 5
                                                    '�ʒu�݌�
            Printer.Print Tab(X_Tab);

            If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If

            Edit = ""
            If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
                Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#,##0")
                If Len(Edit) < 9 Then
                    Edit = Space(9 - Len(Edit)) & Edit
                End If
                Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & _
                       Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & _
                       Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & _
                       Right(EE_ZAIKO_TBL(0).EE_LOC, 2) & Edit
            End If

            Printer.Print Edit

            Printer.Print
        
            LCNT = LCNT + 2
    
        End If
        com = BtOpGetNext
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(LCNT As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If LCNT < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    If LCNT = LCTL Then
    Else
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i

    Printer.Print Tab(MGN_L + 35);
    
    Printer.Print "���i���x���A���[�����X�g(�o�׃f�[�^�Ή�)";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "�q�ɁF";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  "
'    Printer.Print "�i�ݒ蔭���_ " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "���j"
    Printer.Print

'    Printer.Print Tab(MGN_L);
'    Printer.Print "�W���I��";
'    Printer.Print Tab(MGN_L + 13);
'    Printer.Print "�i�ԁi�O���j";
'    Printer.Print Tab(MGN_L + 26);
'    Printer.Print "����(����)";
'    Printer.Print Tab(MGN_L + 38);
'    Printer.Print "���i���ύ݌�";
'    Printer.Print Tab(MGN_L + 58);
'    Printer.Print "�����i�݌�";
'    Printer.Print Tab(MGN_L + 74);
'    Printer.Print "�����Ϗo�א�";
'    Printer.Print Tab(MGN_L + 88);
'    Printer.Print "���O���i���K�v��";
'    Printer.Print Tab(MGN_L + 108);
'    Printer.Print "���O���i����"
'
'    Set Printer.Font = MidFont
'    Printer.Print Tab(MGN_L + 112);
'    Printer.Print "(�ߋ�3����ԕ���)";
'    Printer.Print Tab(MGN_L + 130);
'    Printer.Print "(�����Ϗo�א�-���i���ύ݌�)";
'    Printer.Print Tab(MGN_L + 158);
'    Printer.Print "(���i���ύ݌�/�����Ϗo�א�)"
'
'
'    Set Printer.Font = NormalFont

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "����";
    Printer.Print Tab(MGN_L + 42);
    Printer.Print "���ϐ�";
    Printer.Print Tab(MGN_L + 54);
    Printer.Print "�����i";
    Printer.Print Tab(MGN_L + 62);
    Printer.Print "�{���o�א�";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "�K�v��";
    Printer.Print Tab(MGN_L + 90);
    Printer.Print "�@��";
    Printer.Print Tab(MGN_L + 113);
    Printer.Print "�ʒu�݌�"

    Printer.Print

    LCNT = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   �x���p�W�v�f�[�^�쐬����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer


Dim SKIP_Flg    As Boolean
    
    Data_Make_Proc = True

'---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_POS, GOODSREC, Len(GOODSREC), K1_GOODS, Len(K1_GOODS), 1)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "���i���x���W�v�f�[�^")
                    Exit Function
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            
            sts = BTRV(BtOpDelete, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "���i���x���W�v�f�[�^")
                    Exit Function
            End Select
        
        Loop
        
        com = BtOpGetNext
    
    Loop
'---------------------------------------------------------- '�w�}�[�f�[�^�x�[�X�ō쐬�iKEY�̂݁j

    
    
    Call UniCode_Conv(K3_P_SSHIJI_O.HAKKO_DT, Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text)
    Call UniCode_Conv(K3_P_SSHIJI_O.TORI_KBN, Right(Combo(pcmbTORI_KBN).Text, 1))
    Call UniCode_Conv(K3_P_SSHIJI_O.UKEHARAI_CODE, Text(ptxUKEHARAI_CODE).Text)
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        sts = BTRV(com, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K3_P_SSHIJI_O, Len(K3_P_SSHIJI_O), 3)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_O_REC.HAKKO_DT, vbUnicode) > Text(ptxE_YY).Text & _
                                                                    Text(ptxE_MM).Text & _
                                                                    Text(ptxE_DD).Text Then
                    '���t�͈͊O
                    Exit Do
                End If
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���w�}�\�f�[�^")
                Exit Function
        End Select
        
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
        SKIP_Flg = False
        
        
        If Trim(Right(Combo(pcmbTORI_KBN).Text, 1)) <> "" Then
            If StrConv(P_SSHIJI_O_REC.TORI_KBN, vbUnicode) <> Right(Combo(pcmbTORI_KBN).Text, 1) Then
                '�����敪
                SKIP_Flg = True
            End If
        End If
        
        
        
        
        If Trim(Text(ptxUKEHARAI_CODE).Text) <> "" Then
            If Trim(StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)) <> Trim(Text(ptxUKEHARAI_CODE).Text) Then
                '�����R�[�h�i�󕥐�R�[�h�j
                SKIP_Flg = True
            End If
        End If
        
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
        
        If StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
            StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
            SKIP_Flg = True
        End If
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    
    
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> GOODS_ON Then
                    SKIP_Flg = True
                End If
            Case BtErrKeyNotFound
                SKIP_Flg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Function
        End Select
                
                
        If Trim(Text(ptxSOKO_NO).Text) <> "" Then
            If Text(ptxSOKO_NO).Text <> StrConv(ITEMREC.ST_SOKO, vbUnicode) Then
                SKIP_Flg = True
            End If
        End If
            
            
            
        If Not SKIP_Flg Then
            
            Call UniCode_Conv(K2_GOODS.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K2_GOODS.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K2_GOODS.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            
            sts = BTRV(BtOpGetEqual, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
            Select Case sts
                Case BtNoErr
                    SKIP_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���i���W�v�t�@�C��")
                    Exit Function
            End Select
            
            
            If Not SKIP_Flg Then
If Trim(StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)) = "AXW3482-250" Then
    Debug.Print
End If
        
                                                        '���ƕ�
                Call UniCode_Conv(GOODSREC.JGYOBU, Last_JGYOBU)
                                                        '�����O
                Call UniCode_Conv(GOODSREC.NAIGAI, NAIGAI_NAI)
                                                        '�i�ԁi�O���j
                Call UniCode_Conv(GOODSREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                        '�W���I��
                Call UniCode_Conv(GOODSREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                Call UniCode_Conv(GOODSREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                        '����
                Call UniCode_Conv(GOODSREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
            
            
                Call UniCode_Conv(GOODSREC.Sumi_QTY, "00000000")        '���i���ςݍ݌ɐ�
                Call UniCode_Conv(GOODSREC.Mi_QTY, "00000000")          '�����i�݌ɐ�
                Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")       '���Ϗo�א�
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")    '���O���i����
            
            
                sts = BTRV(BtOpInsert, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
                Select Case sts
                    Case BtNoErr
                    Case Else
                        Call File_Error(sts, BtOpInsert, "���i���v��x��")
                        Exit Function
                End Select
            
            End If
            
        End If
            
        com = BtOpGetNext
            
    Loop
            
            
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���i���W�v�t�@�C��")
                Exit Function
        End Select
Debug.Print StrConv(GOODSREC.HIN_GAI, vbUnicode)
        If Data_Make_Sub() Then
            Exit Function
        End If
            
        
        com = BtOpGetNext
    
    Loop


    Data_Make_Proc = False


End Function

Private Function Data_Proc() As Integer
'----------------------------------------------------------------------------
'                   �b�r�u�f�[�^�쐬����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Save_Soko       As String * 2

Dim Edit            As String

Dim FileNo          As Integer
Dim fileName        As String
    
    
Dim SKIP_Flg        As Boolean
    
    
    Data_Proc = True

    Call Input_Lock

    fileName = GOODS_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - sts)
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo
    On Error GoTo 0


    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If
    
    Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GOODS.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_GOODS.ST_SOKO, "")
    Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
    Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select
'-------------------------------------------------  ���׈��
        
        SKIP_Flg = False
        Select Case Right(Combo(pcmbMI_ZAIKO).Text, 1)
            Case "0"        '�S��
            Case "1"        '0�ΏۊO
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) = 0 Then
                    SKIP_Flg = True
                End If
            Case "2"        '0�̂�
                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <> 0 Then
                    SKIP_Flg = True
                End If
        End Select
        
        If SKIP_Flg Then
        Else
            If com = BtOpGetGreater Then
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        End If
                    Case BtErrKeyNotFound
                        '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                        '�w�b�_�[�o��
                Write #FileNo, "*** ���i���x���A���[�����X�g�@***",
                Write #FileNo, "�쐬���t:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                        
            
                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu��1", "�����i�@�ʒu��2", "�����i�@�ʒu��3", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����"
                
            
                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(�����_" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
            
            
            End If
        
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                        End If
                    
                    Case BtErrKeyNotFound
                            '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                
                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(�����_" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
                
            End If
        
        
            
            
                                                    '�W���I��
                            
            Edit = StrConv(SOKOREC.Soko_No, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
            Write #FileNo, Edit,
                                                    '�i�ԁi�O���j

            Write #FileNo, StrConv(GOODSREC.HIN_GAI, vbUnicode),
                                                    '����
            Write #FileNo, StrConv(GOODSREC.PACKING_NO, vbUnicode),
                                                    '���i���ςݍ݌ɐ�
            Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '�����i�݌ɐ�
            Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    
            If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If
                                                    '�����i�ʒu��
            If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(0).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
            If Len(Trim(EE_ZAIKO_TBL(1).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(1).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(1).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
            If Len(Trim(EE_ZAIKO_TBL(2).EE_LOC)) = 0 Then
                Write #FileNo, ,
            Else
                Edit = Left(EE_ZAIKO_TBL(2).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(2).EE_LOC, 2)
                Edit = Edit & " " & Format(EE_ZAIKO_TBL(2).EE_QTY, "#0")
                Write #FileNo, Edit,
            End If
                                                    
                                                    '�����Ϗo�א�
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '���O���i���K�v��
            Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '���O���i����
            Edit = Format(CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit
                
            com = BtOpGetNext
        End If
    Loop

    Close #FileNo

    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"

    Call Input_UnLock
    
    Data_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Proc = True
    End If


End Function

Private Function MI_ZAIKO_KENSAKU(Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   �����i�̏���
'----------------------------------------------------------------------------
Dim i           As Integer

Dim com         As Integer
Dim sts         As Integer

    MI_ZAIKO_KENSAKU = True
    
    For i = 0 To UBound(EE_ZAIKO_TBL)
        EE_ZAIKO_TBL(i).EE_LOC = ""
        EE_ZAIKO_TBL(i).EE_QTY = 0
    Next i
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI_NAI Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
                    Exit Do
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Exit Function
        End Select
        
        
        If StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
            StrConv(ZAIKOREC.Retu, vbUnicode) & _
            StrConv(ZAIKOREC.Ren, vbUnicode) & _
            StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(GOODSREC.ST_SOKO, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_RETU, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_REN, vbUnicode) & _
                                                    StrConv(GOODSREC.ST_DAN, vbUnicode) Then
        
        
            For i = 0 To UBound(EE_ZAIKO_TBL)
                            
                If Trim(EE_ZAIKO_TBL(i).EE_LOC) = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                    Exit For
                Else
                    If Len(Trim(EE_ZAIKO_TBL(i).EE_LOC)) = 0 Then
                        EE_ZAIKO_TBL(i).EE_LOC = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                        Exit For
                    End If
                End If
            Next i
        
            If i > UBound(EE_ZAIKO_TBL) Then
                Exit Do
            End If
                
        
            EE_ZAIKO_TBL(i).EE_QTY = EE_ZAIKO_TBL(i).EE_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
    
        End If
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function

Public Function F106021_Zaiko_Syukei_Proc(Sumi_Zaiko_Qty As Long, _
                                    Mi_Zaiko_Qty As Long, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional LOCATION As String = "        ") As Integer
'****************************************************
'*      �݌ɐ��W�v
'*
'*  �i�Ԃ܂��͕i�ԁ{�I�Ԗ��̍݌ɐ����W�v����B
'*
'*  ���� :  �݌ɐ��i���i���ς݁j
'*          �݌ɐ��i�����i�j
'*          ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          �I��(�ȗ��� �ȗ�=��)
'*  �߂�l: false    ����
'*          SYS_ERR  �p���ł��Ȃ��ُ�
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Soko_No     As String * 2
Dim Retu        As String * 2
Dim Ren         As String * 2
Dim Dan         As String * 2
    
Dim Not_GOODS   As Boolean

Dim i           As Integer

    F106021_Zaiko_Syukei_Proc = SYS_ERR

    Sumi_Zaiko_Qty = 0
    Mi_Zaiko_Qty = 0

    com = BtOpGetGreater

    If Len(Trim(LOCATION)) = 0 Then
                                '�q�ɔԍ��󔒂͒I�ԏȗ��Ƃ݂Ȃ�
        Call UniCode_Conv(K1_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
        Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
        Call UniCode_Conv(K1_ZAIKO.Retu, "")
        Call UniCode_Conv(K1_ZAIKO.Ren, "")
        Call UniCode_Conv(K1_ZAIKO.Dan, "")

        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop

    Else

        Soko_No = Mid(LOCATION, 1, 2)
        Retu = Mid(LOCATION, 3, 2)
        Ren = Mid(LOCATION, 5, 2)
        Dan = Mid(LOCATION, 7, 2)

        Call UniCode_Conv(K0_ZAIKO.Soko_No, Soko_No)
        Call UniCode_Conv(K0_ZAIKO.Retu, Retu)
        Call UniCode_Conv(K0_ZAIKO.Ren, Ren)
        Call UniCode_Conv(K0_ZAIKO.Dan, Dan)
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        Do
            sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(Retu)) = 0 Then
                        Retu = StrConv(ZAIKOREC.Retu, vbUnicode)
                    End If
                    If Len(Trim(Ren)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Ren, vbUnicode)
                    End If
                    If Len(Trim(Dan)) = 0 Then
                        Ren = StrConv(ZAIKOREC.Dan, vbUnicode)
                    End If

                    If Soko_No <> StrConv(ZAIKOREC.Soko_No, vbUnicode) Or _
                        Retu <> StrConv(ZAIKOREC.Retu, vbUnicode) Or _
                        Ren <> StrConv(ZAIKOREC.Ren, vbUnicode) Or _
                        Dan <> StrConv(ZAIKOREC.Dan, vbUnicode) Or _
                        JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ƀf�[�^")
                    Exit Function
            End Select

            Select Case StrConv(ZAIKOREC.GOODS_ON, vbUnicode)
                Case "0"
                    Not_GOODS = False
                    
                                 
                    For i = 0 To UBound(SHO_SOKO)
                    
                    
                        If StrConv(ZAIKOREC.Soko_No, vbUnicode) = SHO_SOKO(i) Then
                            Not_GOODS = True
                            Exit For
                        End If
                    
                    
                    Next i
                    
                    If Not_GOODS Then
                        Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    Else
                    
                        Sumi_Zaiko_Qty = Sumi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                    End If
                Case "1"
                    Mi_Zaiko_Qty = Mi_Zaiko_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End Select

            com = BtOpGetNext

            DoEvents
        Loop
    End If

    F106021_Zaiko_Syukei_Proc = False

End Function



Private Function Data_Make_Sub() As Integer
    
Dim sts         As Integer
Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
Dim AVE_QTY     As Long
Dim ans         As Integer
    
    
    Data_Make_Sub = True
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))


    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Data_Make_Sub = False
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select

    '-----------------------------------------  '���i���W�v�t�@�C���쐬
                                                '�݌ɏW�v����
    If F106021_Zaiko_Syukei_Proc(Sumi_QTY, _
                            Mi_QTY, _
                            StrConv(ITEMREC.JGYOBU, vbUnicode), _
                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
        Exit Function
    End If
                                                
                                                
    If Sumi_QTY > 0 Then
        Sumi_QTY = Sumi_QTY - 1             '�T���v�����}�C�i�X
    End If
                                            '���i���ςݍ݌ɐ�
    Call UniCode_Conv(GOODSREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                            '�����i�݌ɐ�
    Call UniCode_Conv(GOODSREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                    
                                            '�����Ϗo�א�
    Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
    
    AVE_QTY = 0
    sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    Select Case sts
        Case BtNoErr
            Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
            AVE_QTY = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
        Case BtErrKeyNotFound
            Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�")
            Exit Function
    End Select
    Call UniCode_Conv(GOODSREC.AVE_SYUKA, Format(AVE_QTY, "00000000"))
                                            '���O���i����
    If AVE_QTY = 0 Then
        Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")
    Else
        Call UniCode_Conv(GOODSREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
    End If
        
        
    Do
        
        sts = BTRV(BtOpUpdate, GOODS_POS, GOODSREC, Len(GOODSREC), K2_GOODS, Len(K2_GOODS), 2)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���i���x���W�v�f�[�^")
                Exit Function
        End Select
    
    Loop

    Data_Make_Sub = False


End Function

Private Function Ukeharai_Set_Proc() As Integer

Dim com As Integer
Dim sts As Integer

    
    
    Ukeharai_Set_Proc = True
    
    
    
    Combo(pcmbUKEHARAI_CODE).Clear
    Combo(pcmbUKEHARAI_CODE).AddItem "�S�@�ā@�@�@�@�@�@�@�@�@�@�@�@" & "     " & " "
    
    
    
    
    Call UniCode_Conv(K1_P_UKEHARAI.TORI_KBN, Right(Combo(pcmbTORI_KBN), 1))
    Call UniCode_Conv(K1_P_UKEHARAI.UKEHARAI_CODE, "")


    com = BtOpGetGreaterEqual


    Do
        DoEvents
        
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K1_P_UKEHARAI, Len(K1_P_UKEHARAI), 1)
        Select Case sts
            Case BtNoErr
                If Trim(Right(Combo(pcmbTORI_KBN), 1)) <> "" Then
                    If Right(Combo(pcmbTORI_KBN), 1) <> StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) Then
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        End Select
        
        
        Combo(pcmbUKEHARAI_CODE).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & _
                                            "     " & StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    Loop

    Combo(pcmbUKEHARAI_CODE).ListIndex = 0


    Ukeharai_Set_Proc = False


End Function
