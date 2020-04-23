Attribute VB_Name = "P_CODE_TBL"
Option Explicit
'********************************************************************
'*                                                                  *
'*              �R�[�h�}�X�^  �敪��`      �@                      *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
'�ް��敪���e�e�[�u��
Private Type P_KBN_TBL_Tag
    KBN_CD      As String * 2               '����
    KBN_NM      As String                   '����
    KBN_Len     As Integer                  'KEY�L����
    KBN_OP1     As Integer                  '��߼��1
    KBN_OP2     As Integer                  '��߼��2
    KBN_OP_NM1  As String                   '��߼��1
    KBN_OP_NM2  As String                   '��߼��2
    KBN_BIKOU   As String                   '�ŗLү����
End Type

Public P_KBN_TBL(0 To P_KBN_MAX) As P_KBN_TBL_Tag


Public Sub P_CODE_TBL_Proc()
'********************************************************************
'*                                                                  *
'*              �R�[�h�}�X�^  �敪�ݒ�      �@                      *
'*                                                                  *
'********************************************************************
                                
Dim c   As String * 128
                                
                                '�敪ð��پ��
    P_KBN_TBL(0).KBN_CD = P_KBN01_CD                '�d���敪�@     �R�[�h
    P_KBN_TBL(0).KBN_NM = P_KBN01_NM                '               ����
    P_KBN_TBL(0).KBN_Len = P_KBN01_Len              '               �L����
    P_KBN_TBL(0).KBN_OP1 = P_KBN01_OP1              '               ��߼��1
    P_KBN_TBL(0).KBN_OP2 = P_KBN01_OP2              '               ��߼��2
    P_KBN_TBL(0).KBN_OP_NM1 = P_KBN01_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(0).KBN_OP_NM2 = P_KBN01_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "01", "P_SYS", c) Then
        P_KBN_TBL(0).KBN_BIKOU = ""
    Else
        P_KBN_TBL(0).KBN_BIKOU = Trim(c)
    End If
    
    P_KBN_TBL(1).KBN_CD = P_KBN02_CD                '�̔��敪�@     �R�[�h
    P_KBN_TBL(1).KBN_NM = P_KBN02_NM                '               ����
    P_KBN_TBL(1).KBN_Len = P_KBN02_Len              '               �L����
    P_KBN_TBL(1).KBN_OP1 = P_KBN02_OP1              '               ��߼��1
    P_KBN_TBL(1).KBN_OP2 = P_KBN02_OP2              '               ��߼��2
    P_KBN_TBL(1).KBN_OP_NM1 = P_KBN02_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(1).KBN_OP_NM2 = P_KBN02_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "02", "P_SYS", c) Then
        P_KBN_TBL(1).KBN_BIKOU = ""
    Else
        P_KBN_TBL(1).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(2).KBN_CD = P_KBN03_CD                '���x�P�ʁ@     �R�[�h
    P_KBN_TBL(2).KBN_NM = P_KBN03_NM                '               ����
    P_KBN_TBL(2).KBN_Len = P_KBN03_Len              '               �L����
    P_KBN_TBL(2).KBN_OP1 = P_KBN03_OP1              '               ��߼��1
    P_KBN_TBL(2).KBN_OP2 = P_KBN03_OP2              '               ��߼��2
    P_KBN_TBL(2).KBN_OP_NM1 = P_KBN03_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(2).KBN_OP_NM2 = P_KBN03_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "03", "P_SYS", c) Then
        P_KBN_TBL(2).KBN_BIKOU = ""
    Else
        P_KBN_TBL(2).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(3).KBN_CD = P_KBN04_CD                '�d������@     �R�[�h
    P_KBN_TBL(3).KBN_NM = P_KBN04_NM                '               ����
    P_KBN_TBL(3).KBN_Len = P_KBN04_Len              '               �L����
    P_KBN_TBL(3).KBN_OP1 = P_KBN04_OP1              '               ��߼��1
    P_KBN_TBL(3).KBN_OP2 = P_KBN04_OP2              '               ��߼��2
    P_KBN_TBL(3).KBN_OP_NM1 = P_KBN04_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(3).KBN_OP_NM2 = P_KBN04_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "04", "P_SYS", c) Then
        P_KBN_TBL(3).KBN_BIKOU = ""
    Else
        P_KBN_TBL(3).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(4).KBN_CD = P_KBN05_CD                '���P�^�S����   �R�[�h
    P_KBN_TBL(4).KBN_NM = P_KBN05_NM                '               ����
    P_KBN_TBL(4).KBN_Len = P_KBN05_Len              '               �L����
    P_KBN_TBL(4).KBN_OP1 = P_KBN05_OP1              '               ��߼��1
    P_KBN_TBL(4).KBN_OP2 = P_KBN05_OP2              '               ��߼��2
    P_KBN_TBL(4).KBN_OP_NM1 = P_KBN05_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(4).KBN_OP_NM2 = P_KBN05_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "05", "P_SYS", c) Then
        P_KBN_TBL(4).KBN_BIKOU = ""
    Else
        P_KBN_TBL(4).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(5).KBN_CD = P_KBN06_CD                '����           �R�[�h
    P_KBN_TBL(5).KBN_NM = P_KBN06_NM                '               ����
    P_KBN_TBL(5).KBN_Len = P_KBN06_Len              '               �L����
    P_KBN_TBL(5).KBN_OP1 = P_KBN06_OP1              '               ��߼��1
    P_KBN_TBL(5).KBN_OP2 = P_KBN06_OP2              '               ��߼��2
    P_KBN_TBL(5).KBN_OP_NM1 = P_KBN06_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(5).KBN_OP_NM2 = P_KBN06_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "06", "P_SYS", c) Then
        P_KBN_TBL(5).KBN_BIKOU = ""
    Else
        P_KBN_TBL(5).KBN_BIKOU = Trim(c)
    End If
   
    P_KBN_TBL(6).KBN_CD = P_KBN07_CD                '��Ж��^���ƕ� �R�[�h
    P_KBN_TBL(6).KBN_NM = P_KBN07_NM                '               ����
    P_KBN_TBL(6).KBN_Len = P_KBN07_Len              '               �L����
    P_KBN_TBL(6).KBN_OP1 = P_KBN07_OP1              '               ��߼��1
    P_KBN_TBL(6).KBN_OP2 = P_KBN07_OP2              '               ��߼��2
    P_KBN_TBL(6).KBN_OP_NM1 = P_KBN07_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(6).KBN_OP_NM2 = P_KBN07_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "07", "P_SYS", c) Then
        P_KBN_TBL(6).KBN_BIKOU = ""
    Else
        P_KBN_TBL(6).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(7).KBN_CD = P_KBN08_CD                '���ދ敪
    P_KBN_TBL(7).KBN_NM = P_KBN08_NM                '               ����
    P_KBN_TBL(7).KBN_Len = P_KBN08_Len              '               �L����
    P_KBN_TBL(7).KBN_OP1 = P_KBN08_OP1              '               ��߼��1
    P_KBN_TBL(7).KBN_OP2 = P_KBN08_OP2              '               ��߼��2
    P_KBN_TBL(7).KBN_OP_NM1 = P_KBN08_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(7).KBN_OP_NM2 = P_KBN08_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "08", "P_SYS", c) Then
        P_KBN_TBL(7).KBN_BIKOU = ""
    Else
        P_KBN_TBL(7).KBN_BIKOU = Trim(c)
    End If

    '---------------------------------------------  2008.02.28 ��
        
    P_KBN_TBL(8).KBN_CD = P_KBN09_CD                '�o�c����
    P_KBN_TBL(8).KBN_NM = P_KBN09_NM                '               ����
    P_KBN_TBL(8).KBN_Len = P_KBN09_Len              '               �L����
    P_KBN_TBL(8).KBN_OP1 = P_KBN09_OP1              '               ��߼��1
    P_KBN_TBL(8).KBN_OP2 = P_KBN09_OP2              '               ��߼��2
    P_KBN_TBL(8).KBN_OP_NM1 = P_KBN09_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(8).KBN_OP_NM2 = P_KBN09_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "09", "P_SYS", c) Then
        P_KBN_TBL(8).KBN_BIKOU = ""
    Else
        P_KBN_TBL(8).KBN_BIKOU = Trim(c)
    End If

    P_KBN_TBL(9).KBN_CD = P_KBN10_CD                '����
    P_KBN_TBL(9).KBN_NM = P_KBN10_NM                '               ����
    P_KBN_TBL(9).KBN_Len = P_KBN10_Len              '               �L����
    P_KBN_TBL(9).KBN_OP1 = P_KBN10_OP1               '               ��߼��1
    P_KBN_TBL(9).KBN_OP2 = P_KBN10_OP2              '               ��߼��2
    P_KBN_TBL(9).KBN_OP_NM1 = P_KBN10_OP_NM1        '               ��߼�ݖ���1
    P_KBN_TBL(9).KBN_OP_NM2 = P_KBN10_OP_NM2        '               ��߼�ݖ���2
                                                    '�ŗLү���ގ�荞��
    If GetIni(StrConv(App.EXEName, vbUpperCase), "10", "P_SYS", c) Then
        P_KBN_TBL(9).KBN_BIKOU = ""
    Else
        P_KBN_TBL(9).KBN_BIKOU = Trim(c)
    End If

End Sub



Public Sub P_CODE_INI_TBL_Proc()
'********************************************************************
'*                                                                  *
'*              �R�[�h�}�X�^  �敪�ݒ�      �@                      *
'*                                                                  *
'*  2018.04.07  INI-->TBL                                           *
'********************************************************************
                                
Dim c   As String * 128
Dim i   As Integer
                                
                                
    For i = 0 To 9
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_CD", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_CD = ""
        Else
            P_KBN_TBL(i).KBN_CD = Trim(c)
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_NM", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_NM = ""
        Else
            P_KBN_TBL(i).KBN_NM = Trim(c)
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_Len", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_Len = 0
        Else
            P_KBN_TBL(i).KBN_Len = Val(Trim(c))
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_OP1", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_OP1 = 0
        Else
            P_KBN_TBL(i).KBN_OP1 = Val(Trim(c))
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_OP2", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_OP2 = 0
        Else
            P_KBN_TBL(i).KBN_OP2 = Val(Trim(c))
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_OP_NM1", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_OP_NM1 = ""
        Else
            P_KBN_TBL(i).KBN_OP_NM1 = Trim(c)
        End If
    
        If GetIni(App.EXEName, "P_KBN" & Format(i + 1, "00") & "_OP_NM2", App.EXEName, c) Then
            P_KBN_TBL(i).KBN_OP_NM2 = ""
        Else
            P_KBN_TBL(i).KBN_OP_NM2 = Trim(c)
        End If
    
        If GetIni(StrConv(App.EXEName, vbUpperCase), Format(i + 1, "00"), App.EXEName, c) Then
            P_KBN_TBL(0).KBN_BIKOU = ""
        Else
            P_KBN_TBL(0).KBN_BIKOU = Trim(c)
        End If
    
    
    Next i
                                
End Sub

