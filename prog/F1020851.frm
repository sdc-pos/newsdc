VERSION 5.00
Begin VB.Form F1020851 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���Ƀ��X�g�f�[�^�o�́i�I���������j ([F102085] 2011.07.14 11:15)"
   ClientHeight    =   3744
   ClientLeft      =   2028
   ClientTop       =   2268
   ClientWidth     =   8028
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3744
   ScaleWidth      =   8028
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���Ƀ��X�g�p�f�[�^�쐬��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   5760
   End
End
Attribute VB_Name = "F1020851"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Y_NYU_DATA  As String                   '���׃f�[�^�t���p�X

Private Sub Form_Activate()

Dim i               As Integer
    
Dim FileNo          As Integer
Dim fileName        As String
    
Dim Den_Date        As String * 8
Dim Rec_Cnt         As Integer
    
Dim Fast_Flg        As Boolean
    
Dim Ret             As Integer
    
    
    Fast_Flg = True
    For i = 0 To UBound(JGYOBU_T)
    
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Exit For
        End If
        
        
        Last_JGYOBU = JGYOBU_T(i).CODE
    
    
        If Fast_Flg Then
            FileNo = FreeFile
            fileName = Y_NYU_DATA
        
        
            Ret = InStr(1, Trim(fileName), ".") - 1
            fileName = Left(Trim(fileName), Ret) & Right(Format(Now, "YYYY/MM/DD"), 2) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
            
            On Error GoTo Error_Proc
            
            Open (fileName) For Output As FileNo
        
            Rec_Cnt = 0
        End If
    
    
        If Output_Proc(FileNo, Fast_Flg, Rec_Cnt, Den_Date) Then
            Unload Me
        End If
    
    Next i
    
    
    Write #FileNo, "�`�[���t�F", Left(Den_Date, 4) & "/" & Mid(Den_Date, 5, 2) & "/" & Right(Den_Date, 2), "�f�[�^�����F", Format(Rec_Cnt, "#0")


    Close #FileNo
    
    
    Unload Me

Error_Proc:

    If Err.Number = 70 Then
        Call LOG_OUT(LOG_F, "�u" & fileName & "�v" & "�g�p���I�I�@�u���Ƀ��X�g�v�f�[�^�o�͂ł��܂���ł����B")
    Else
        Call LOG_OUT(LOG_F, "Err.Number=" & Err.Number & "�u���Ƀ��X�g�v�f�[�^�o�͂ł��܂���ł����B")
    End If

    Unload Me

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)


                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If


                                '���Ƀ��X�g�f�[�^�t�@�C������荞��
    If GetIni("FILE", "Y_NYU_DATA", "SYS", c) Then
        Beep
        MsgBox "���Ƀ��X�g�f�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Y_NYU_DATA = Trim(c)
                                '���׃f�[�^�t�@�C��OPEN
    If Y_NYU_Open(0) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C��OPEN
    If ZAIKO_Open(0) Then
        Unload Me
    End If
                                '�����σf�[�^�t�@�C��OPEN
    If AVE_SYUKA_Open(0) Then
        Unload Me
    End If
    
    

    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '���׃f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���׃f�[�^�t�@�C��")
        End If
    End If
                                            '�݌Ƀf�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^�t�@�C��")
        End If
    End If
                                            '�����Ϗo�ׂb�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo��")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020851 = Nothing

    End
End Sub



Private Function Output_Proc(FileNo As Integer, Fast_Flg As Boolean, Rec_Cnt As Integer, Den_Date As String) As Integer
'----------------------------------------------------------------------------
'                   �b�r�u�f�[�^�o�͏���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Ret             As Integer

Dim c               As String * 128

Dim i               As Integer


Dim Skip_Flg        As Boolean

Dim Work_Soko       As String * 2
Dim Soko_No         As String * 2

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long


    Output_Proc = True
'���s���̓C�x���g�擾�s��


    
    com = BtOpGetFirst


    Do
        DoEvents
        
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)

        Select Case sts
            Case BtNoErr
                '�Ώۑq�ɂ̔���
                Skip_Flg = True
                
                If StrConv(Y_NYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Else
                    
                    Select Case Last_JGYOBU
                        Case SOJIKI                     '����
                                                
                            If StrConv(Y_NYUREC.NYU_LIST_OUT, vbUnicode) = "9" Then
                            Else
                                Select Case Trim(StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode))
                                
                                    Case "91H"
                                        Work_Soko = "90"
                                    Case Else
                                        Work_Soko = "80"
                                End Select
    
    '                        If Work_Soko = "90" Then
                                    Skip_Flg = False
    '                        End If
    
    '                        For i = 0 To 1
    '                            If Check1(i).Value Then
    '                                If Trim(Check1(i).Caption) = Work_Soko Then
    '                                    Skip_Flg = False
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next i
                            End If
                        
                        
                        Case DENKA, SUIHAN, SENTAKU     '����
    
                            If StrConv(Y_NYUREC.NYU_LIST_OUT, vbUnicode) = "9" Then
                            Else
                                Skip_Flg = False
                                Select Case Trim(StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode))
                                    Case "G22"
                                        Work_Soko = "80"
                                    Case "G11"
                                        Work_Soko = "81"
                                    Case Else
                                        Work_Soko = "90"
                                End Select
                            End If
    
    
    '
    '                        Select Case Trim(StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode))
    '                            Case "G22"
    '                                Work_Soko = "80"
    '                            Case "G11"
    '                                Work_Soko = "81"
    '                            Case Else
    '                                Work_Soko = "90"
    '                        End Select
    '
    '
    '                        For i = 2 To 4
    '                            If Check1(i).Value Then
    '                                If Trim(Check1(i).Caption) = Work_Soko Then
    '                                    Skip_Flg = False
    '                                    Exit For
    '                                End If
    '                            End If
    '                        Next i
    '
                    
                        Case SENTAKU                     '�܈�
                            
                            
                            If StrConv(Y_NYUREC.NYU_LIST_OUT, vbUnicode) = "9" Then
                            Else
                                Work_Soko = "90"
                                Skip_Flg = False
                            End If
                        
                        Case AIRCON, REIZOU             '����   2007.05.24
                            
                            If StrConv(Y_NYUREC.NYU_LIST_OUT, vbUnicode) = "9" Then
                            Else
                            
    '                            If Format(Now, "YYYYMMDD") <> StrConv(Y_NYUREC.SYUKO_YMD, vbUnicode) Then
    '                            Else
                                    Select Case StrConv(Y_NYUREC.H_SOKO, vbUnicode)
                                    
                                        Case "S8"
                                            Work_Soko = "80"
                                        Case Else
                                            Work_Soko = "90"
                                    End Select
                                    Skip_Flg = False
    '                            End If
                            End If
                    End Select
                                    
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ח\��f�[�^")
                Exit Function
        End Select
    
        If Not Skip_Flg Then
            Rec_Cnt = Rec_Cnt + 1

            If Fast_Flg Then
                Den_Date = StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode)
                Write #FileNo, , , "���Ƀ��X�g", , "�쐬���F", Left(Den_Date, 4) & "/" & Mid(Den_Date, 5, 2) & "/" & Right(Den_Date, 2) & "��"
                Write #FileNo, "�o�͑q��", "�W���I��", "�i�ԁi�O���j", "�i�ԁi�����j", "�`�[��", "���ɐ�", "���ɐ��|�O�ؐ�", "�\�Z�P��", "�׎p", "���ɐ�", "�����i", "���i���ς�", "������", "���ƕ�"
                Fast_Flg = False
            End If

            Write #FileNo, Work_Soko,
            '�W���I��
            
                        
            If GetIni("SOKO_NO", Left(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 2), "SYS", c) Then
                Soko_No = Left(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 2)
            Else
                Soko_No = Trim(c)
            End If
                    
            Write #FileNo, Soko_No & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 3, 2) & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 5, 2) & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 7, 2),
            Write #FileNo, StrConv(Y_NYUREC.HIN_NO, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.HIN_NAI, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.DEN_NO, vbUnicode),
            Write #FileNo, Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)), "#0"),
            Write #FileNo, Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode)), "#0"),
            Write #FileNo, StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode) & " " & StrConv(Y_NYUREC.YOSAN_TO, vbUnicode), , ,
                    
        
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(Y_NYUREC.JGYOBU, vbUnicode), _
                                    StrConv(Y_NYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_NYUREC.HIN_NO, vbUnicode)) Then
                Exit Function
            End If
        
            Write #FileNo, Format(MI_QTY, "#0"),
            Write #FileNo, Format(SUMI_QTY, "#0"),
        
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            
            Select Case sts
                Case BtNoErr
                    Write #FileNo, Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#0"), Last_JGYOBU
                Case BtErrKeyNotFound
                    Write #FileNo, , Last_JGYOBU
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�")
                    Exit Function
            End Select
        
        End If

        com = BtOpGetNext
    Loop



    
    
    
    
    
    Output_Proc = False


End Function
