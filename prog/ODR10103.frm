VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR10103 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�q���i�W�J��� <2010.06.15>"
   ClientHeight    =   8610
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   7935
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   4020
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   12726
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "�q�i��"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "�K�v����"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "�j�d�x����"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "�s����"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "�\��"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4075"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3942"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2566"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2434"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8706"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2566"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2434"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2566"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2434"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=�l�r �o�S�V�b�N"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "�q���i�@�W�J�^�s�����"
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      DeadAreaBackColor=   -2147483643
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=�l�r �o�S�V�b�N"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFF00&,.bold=0,.fontsize=1125"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=�l�r �S�V�b�N"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=�l�r �S�V�b�N"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=�l�r �S�V�b�N"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2,.alignment=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=16,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=138,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=135,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=136,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=137,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=24,.parent=87,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=21,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=22,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=23,.parent=91"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "ODR10103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�R���{�p�Y��
'Private Const pcmbHBUN = 0



'�O���b�h�p��`
Private ORDR_GRID   As New XArrayDB


Private Const Min_Row% = 1              '�ŏ��s��
Private Const Max_Row = 9999            '�ő�s��


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 4              '�ő��

Private Const Col_ITEM% = 0             '���i�R�[�h
Private Const Col_QTY% = 1              '�g�p����
Private Const Col_Req% = 3              '�K�v����
Private Const Col_OK_DT% = 4            '�\��

Dim row         As Long                 '�Ώہ@�s

Dim Cor_Row     As Long                 '�J�����g�s

Dim Init_F      As Integer
Private Function Data_Disp() As Integer
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim svJGYOBU    As String * 1
Dim svNAIGAI    As String * 1
Dim svHin_gai   As String * 20

Dim sumQty      As Double
Dim sumReq      As Double

Dim W_SeqKey    As Integer
Dim W_STR       As String


    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '��ʍ��ڃ��b�N
    
    DoEvents
    
    Set ORDR_GRID = Nothing
    
    row = 0
    
    Call UniCode_Conv(K0_ODR_REQ.SHIMUKE, GW_SIMUKE)       '�d������
    Call UniCode_Conv(K0_ODR_REQ.JGYOBU, GW_JIGYOBU)       '���ƕ�
    Call UniCode_Conv(K0_ODR_REQ.NAIGAI, GW_NAIGAI)        '�����O
    Call UniCode_Conv(K0_ODR_REQ.HIN_GAI, GW_HINGAI)       '�e�i��
    Call UniCode_Conv(K0_ODR_REQ.ORDER_NO, DIS_ORDR_NO)    '�e�i�ԁ@������
    Call UniCode_Conv(K0_ODR_REQ.INS_NO, DIS_KEY)          '�o�^��
    Call UniCode_Conv(K0_ODR_REQ.BUN_NO, DIS_BUNNO)                '���[��
    
    Call UniCode_Conv(K0_ODR_REQ.KO_HIN_GAI, "")            '�q�i��
    
    svJGYOBU = ""
    svNAIGAI = ""
    svHin_gai = ""
    sumQty = 0
    sumReq = 0
    
    com = BtOpGetGreaterEqual
        
    Do
        sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                Exit Do
                        
            Case Else
                Call File_Error(sts, com, "ODR_REQ")
                Exit Function
        End Select
        
        If Trim(StrConv(ODR_REQ_R.SHIMUKE, vbUnicode)) <> Trim(GW_SIMUKE) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI) Then
            Exit Do
        End If
               
        If Trim(StrConv(ODR_REQ_R.ORDER_NO, vbUnicode)) <> Trim(DIS_ORDR_NO) Then
            Exit Do
        End If
            
        If Trim(StrConv(ODR_REQ_R.INS_NO, vbUnicode)) <> Trim(DIS_KEY) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.BUN_NO, vbUnicode)) <> Trim(DIS_BUNNO) Then
            Exit Do
        End If
            
        Key_Ko_JIGYOBU = Trim(StrConv(ODR_REQ_R.KO_JGYOBU, vbUnicode))
        Key_Ko_NAIGAI = Trim(StrConv(ODR_REQ_R.KO_NAIGAI, vbUnicode))
                '�ҏW
        row = row + 1
        If Grid_Set_Proc() Then
            Exit Function
        End If
        
        com = BtOpGetNext
    Loop
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    '                                            2010/06/15�ǉ�
    '                                   �W�J���̉��Ɂu�\�����v��\������B
    
        '   �ŏ��Ɂu�e���R�[�h�v���擾�B
    Do
        
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, GW_SIMUKE)
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, GW_JIGYOBU)
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, GW_NAIGAI)
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, GW_HINGAI)
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
        com = BtOpGetGreaterEqual
        sts = BTRV(com, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    'Beep
                    'MsgBox "�w�肳�ꂽ�H��������܂���B"
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                yn = MsgBox("���Ŏg�p���ł��I<�\���e>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                If yn = vbNo Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "P_COMPO")
                Exit Function
        End Select
        If sts <> BtNoErr Then
            Exit Do
        End If
        
    
        '   ��������u�q���i���R�[�h�v��ǂ݂Ȃ���W�J�e���o�͂���B
        com = BtOpGetNext
        Do
            sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound, BtErrEOF      '���R�[�h����
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                    yn = MsgBox("���Ŏg�p���ł��I<�\���e>" & Chr(13) & Chr(10) & _
                                "�@�Ď��s���܂����H", vbYesNo + vbExclamation, "�m�F����")
                    If yn = vbNo Then Exit Function
                Case Else
                    Call File_Error(sts, com, "P_COMPO")
                    Exit Function
            End Select
            If sts <> BtNoErr Then
                Exit Do
            End If
            If Trim(StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode)) <> Trim(GW_SIMUKE) Then Exit Do
            If Trim(StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)) <> Trim(GW_JIGYOBU) Then Exit Do
            If Trim(StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode)) <> Trim(GW_NAIGAI) Then Exit Do
            If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(GW_HINGAI) Then Exit Do
            'If Trim(StrConv(P_COMPO_O_REC.DATA_KBN, vbUnicode)) <> "0" Then Exit Do
        
            W_SeqKey = StrConv(P_COMPO_K_REC.SEQNO, vbUnicode)
        
            If CInt(W_SeqKey) <> 0 Then                 '�\�����i���R�[�h�H
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                Call UniCode_Conv(K0_ITEM.NAIGAI, GW_NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)))
                    
                Do
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            
                    Select Case sts
                        Case BtNoErr
                            
                            Exit Do
                            
                        Case BtErrKeyNotFound, BtErrEOF
                            Exit Do
                            
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If yn <> vbYes Then
                                Exit Do
                            End If
                            
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Do
                    End Select
                Loop
                
                If sts <> BtNoErr Then
                            '�ҏW
                    row = row + 1
                    
''-----------------------------------'���ƕ��敪
'Public Const SOJIKI$ = "7"          '�|���@
'Public Const DENKA$ = "D"           '�d������
'Public Const SUIHAN$ = "4"          '���ъ�
'Public Const SENTAKU$ = "1"         '����@�i�A�C�����j
'Public Const AIRCON$ = "A"          '�G�A�R��           2004.12.01
'
'Public Const REIZOU$ = "R"          '�①��             2007.05.24
'
'Public Const SETSUBI$ = "B"         '�ݔ�   2007.03.28
'
'Public Const SHIZAI$ = "S"          '����   2005.11.16
'
'
'Public Const JGYOBU_NON$ = "0"      '���ƕ��敪�Ȃ�
                    
                    
                    Select Case StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)
                        Case SETSUBI
                            W_STR = "���ƕ�:�ݔ�"
                        
                        Case Else
                            W_STR = "���ƕ�:" & StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode)
                    End Select
                    
                    
                    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
                    '�i��
                    ORDR_GRID(row, Col_ITEM) = RTrim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            
                    '�W�J��
                    ORDR_GRID(row, Col_QTY) = ""
            
                    '�s����
                    ORDR_GRID(row, Col_Req) = ""
            
                    '�\��
                    ORDR_GRID(row, Col_OK_DT) = W_STR
                
                End If
            End If
        Loop
        Exit Do
    Loop
    
    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<  2010/06/15�����܂�
    
    Set TDBGrid1.Array = ORDR_GRID
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    DoEvents
    
    Call Input_UnLock                             '��ʍ��ڃ��b�N
    
    Data_Disp = False
    
    
End Function



Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �O���b�h�\���i�ړ����f�[�^���e�j
'               Row   �s��
'               mode�@FALSE:����OFF  TRUE:����ON
'----------------------------------------------------------------------------
Dim W_QTY       As Double
Dim W_STR       As String


    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    '�i��
    ORDR_GRID(row, Col_ITEM) = Trim(StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode))
    
    '�W�J��
    W_QTY = CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode))
    ORDR_GRID(row, Col_QTY) = Format(W_QTY, "###,##0.00")
    
    
    
    '�s����
                                                            '09/03/11 0�̎��A�󔒂ɂ��Ă݂��B(^_^;)
    W_QTY = CDbl(StrConv(ODR_REQ_R.FUSOKU_QTY, vbUnicode))
    If W_QTY <> 0 Then
        ORDR_GRID(row, Col_Req) = Format(W_QTY, "###,##0.00")
    Else
        ORDR_GRID(row, Col_Req) = ""
    End If
    
    
    
    '�\��
    If Trim(StrConv(ODR_REQ_R.OK_DT, vbUnicode)) = "" Then
        W_STR = ""
    Else
        W_STR = Mid(StrConv(ODR_REQ_R.OK_DT, vbUnicode), 3, 2) & "/" & _
                Mid(StrConv(ODR_REQ_R.OK_DT, vbUnicode), 5, 2) & "/" _
                    & Right(StrConv(ODR_REQ_R.OK_DT, vbUnicode), 2)
    End If
    ORDR_GRID(row, Col_OK_DT) = W_STR
    
    Grid_Set_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    ODR10103.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR10103)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR10103)


    ODR10103.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
    
            
        Case 0
            Init_F = 0
            Set ORDR_GRID = Nothing
            Set TDBGrid1.Array = ORDR_GRID
            TDBGrid1.ReBind
            TDBGrid1.Update
            DoEvents
            
            ODR10102_Return = True                '�m�F��ʷ�ݾُI��
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR10103.Top = ODR10101.Top + (ODR10101.Height - ODR10103.Height)
    ODR10103.Left = ODR10101.Left + (ODR10101.Width - ODR10103.Width) / 2
    
    
    
    If Data_Disp Then
        Call Input_UnLock                             '��ʍ��ڃ��b�N
    End If
    
    ODR10102_Return = True
    TDBGrid1.SetFocus
    
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()

    Init_F = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Me.Visible = False
    
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    
    'TDBGrid1.Bookmark = -1
    DoEvents
    
End Sub


Private Sub TDBGrid1_DblClick()
                                    '   2010/05/07�ǉ�
Dim W_ORDR  As String
Dim W_STR   As String

Dim W_KO_ITEM   As String

    If IsNull(TDBGrid1.Bookmark) Then Exit Sub
    
    If TDBGrid1.Bookmark = -1 Then Exit Sub
    
    
    Set ORDR_GRID = TDBGrid1.Array
                
        
        '           ���[�w����ʂɈڍs�I
    'Key_SIMUKE = GW_SIMUKE
    'Key_JIGYOBU = GW_JIGYOBU
    'Key_NAIGAI = GW_NAIGAI
        
    Key_Ko_HinGai = Trim(ORDR_GRID(TDBGrid1.Bookmark, Col_ITEM%))
        
    'Key_Ko_HinGai = "C081"
    
    DoEvents
    
    'MsgBox "�q�i�� <" & Key_Ko_HinGai & ">"
    
    ODR10104.Show vbModal

End Sub

