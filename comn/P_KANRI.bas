Attribute VB_Name = "P_KANRI"
Option Explicit
'********************************************************************
'*                                                                  *
'*              Ç}X^  t@Cè`                            *
'*                                                                  *
'*          CREATE 2005.11.11                                       *
'********************************************************************
't@Chc
Public Const P_KANRI_ID$ = "P_KANRI"

'y[WTCY
Private Const P_KANRI_PG_SIZ% = 512

'|WVEubN
Public P_KANRI_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           \¢Ìè`                             *
'*                                                                  *
'********************************************************************
'*************************** Ú¼è` *****************************
'R[hè`
Public Type P_KANRIREC_Tag
    REC_NO(0 To 1)          As Byte         'Úº°ÄÞ
    SHIME_DD(0 To 1)        As Byte         'SDC÷ßú
    xSASHIZU_NO(0 To 4)     As Byte         'w}[(»Ýl+1) ¢gpÆ·é 2007.11.28
    ORDER_NO(0 To 4)        As Byte         '­(»Ýl+1)
    URIAGE_NO(0 To 4)       As Byte         'ÞãÚº°ÄÞ(»Ýl+1)
    
    ZEI_CHANGE_YMD(0 To 7)  As Byte         'ÁïÅÏXút
    NOW_ZEI_RITU(0 To 3)    As Byte         '»@ÁïÅ¦
    NOW_MARUME(0 To 0)      As Byte         '    Ûß
    NEW_ZEI_RITU(0 To 3)    As Byte         'V@ÁïÅ¦
    NEW_MARUME(0 To 0)      As Byte         '    Ûß
    
    SHONIN_CODE(0 To 4)     As Byte         '³FÒº°ÄÞ
    KAISHA_NAME(0 To 29)    As Byte         'ïÐ¼
    CENTER_NAME(0 To 29)    As Byte         'Z^[¼
    TEL_NO(0 To 14)         As Byte         'dbÔ
    FAX_NO(0 To 14)         As Byte         'FAXÔ
    
    URI_MARUME(0 To 0)      As Byte         'ãàzÛß
    SHI_MARUME(0 To 0)      As Byte         'düàzÛß
    
    SASHIZU_NO(0 To 7)      As Byte         '­(»Ýl+1)   2007.11.28
    
    
    NYUKO_S_RATE(0 To 6)    As Byte         'üÉ@ª[g     2008.02.13
    NYUKO_R_RATE(0 To 6)    As Byte         'üÉ@]T¦       2008.02.13
    
    SYUKO_S_RATE(0 To 6)    As Byte         'oÉ@ª[g     2008.02.13
    SYUKO_R_RATE(0 To 6)    As Byte         'oÉ@]T¦       2008.02.13
    
    SYUKA_S_RATE(0 To 6)    As Byte         'oÉ@ª[g     2008.02.13
    SYUKA_R_RATE(0 To 6)    As Byte         'oÉ@]T¦       2008.02.13
    
    KOUTEI_LOT(0 To 5)      As Byte         'Hö@OãHöWbg   2008.02.13
    KOUTEI_S_RATE(0 To 6)   As Byte         'Hö@ª[g             2008.02.13
    KOUTEI_R_RATE(0 To 6)   As Byte         'Hö@]T¦               2008.02.13
    KOUTEI_SHIZAI(0 To 2)   As Byte         'Hö@ÞmF_       2008.02.13
    KOUTEI_BUHIN(0 To 2)    As Byte         'Hö@¯«imF_     2008.02.13
    KOUTEI_LABEL(0 To 2)    As Byte         'Hö@x\t       2008.02.13
    
    MITSUMORI_NO(0 To 7)    As Byte         '©Ï   2008.02.13
    SEIKYU_NO(0 To 7)       As Byte         '¿   2008.02.13
        
    
    MIN_URIAGE_NO(0 To 7)   As Byte         '~j}ã     2008.02.13
    
    
    FILLER(0 To 18)         As Byte         'FILLER
End Type
'f[^Eobt@
Public P_KANRIREC           As P_KANRIREC_Tag




Private Type P_KOTEI_Tag                    '2008.02.13
    KOTEI(0 To 2)       As Byte
End Type

Public Type P_KANRIREC02_Tag                '2008.02.13
    REC_NO(0 To 1)          As Byte         'Úº°ÄÞ
        
    BEF_KOTEI(0 To 9)       As P_KOTEI_Tag  'OHö
    MAIN_KOTEI(0 To 9)      As P_KOTEI_Tag  'ìÆHö
    AFT_KOTEI(0 To 9)       As P_KOTEI_Tag  'ãHö
        
    FUTAI_KOTEI(0 To 4)     As P_KOTEI_Tag  'tÑHö@(»Ý¢gp)
    KEIHI(0 To 4)           As P_KOTEI_Tag  'oï@(»Ý¢gp)
    
    FILLER(0 To 133)        As Byte         'FILLER
End Type
'f[^Eobt@
Public P_KANRIREC02         As P_KANRIREC02_Tag



'L[è`

Type KEY0_P_KANRI           'jdxO
    REC_NO(0 To 1)          As Byte         'Úº°ÄÞ
End Type
    
'L[Ef[^
Public K0_P_KANRI           As KEY0_P_KANRI

Type P_KANRI_FSpeck
    fs                      As BtFileSpeck  ' Ì§²Ù ½Íß¯¸\¢Ì
    ks0                     As BtKeySpeck   ' ·° ½Íß¯¸\¢Ì
End Type

Private P_KANRI_Speck       As P_KANRI_FSpeck
Private Function P_KANRI_Create() As Integer
'********************************************************************
'*                                                                  *
'*              Ç}X^  bqd`sd                            *
'*                                                                  *
'*      ø  :Èµ                                                 *
'*      ßèl:false ³í                                           *
'*             true  Ùí                                           *
'*                                                                  *
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

    P_KANRI_Create = True
                                            'Ç}X^tpXæÝ
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]ÇÝÝG[")
        Exit Function
    End If

    FullPath = RTrim(c)

    P_KANRI_Speck.fs.recoleng = Len(P_KANRIREC)            ' R[h·
    P_KANRI_Speck.fs.PageSize = P_KANRI_PG_SIZ          ' y[WTCY
    P_KANRI_Speck.fs.idexnumb = 1                       ' CfbNX
    P_KANRI_Speck.fs.fileflag = 0                       ' t@CtO
    P_KANRI_Speck.fs.reserve = &H0                      ' \ñÏÝ
    '--------------------------------------------------- L[O ¤
    P_KANRI_Speck.ks0.keypos = 1                        ' L[|WV
    P_KANRI_Speck.ks0.keyleng = 2                       ' L[·
    P_KANRI_Speck.ks0.keyflag = BtKfExt                 ' L[tO
    P_KANRI_Speck.ks0.keytype = Chr(BtKtString)         ' L[^Cv
    P_KANRI_Speck.ks0.reserve = &H0                     ' \ñÏÝ
    '--------------------------------------------------- L[O ¢
    sts = BTRV(BtOpCreate, P_KANRI_POS, P_KANRI_Speck, Len(P_KANRI_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "Ç}X^")
        Exit Function
    End If
    
    P_KANRI_Create = False

End Function

Public Function P_KANRI_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              Ç}X^  nodm
'*
'*      ø  :Open Mode(BtrieveQÆ)
'*      ßèl:false ³í
'*             true  Ùí
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

    P_KANRI_Open = True
                                            'Ç}X^tpXæÝ
    sts = GetIni("FILE", P_KANRI_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [P_KANRI]ÇÝÝG[")
        Exit Function
    End If
    FullPath = RTrim(c)

    Do
        sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_KANRI_Create()      'Ç}X^ì¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "Ç}X^")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "Ç}X^")
                Exit Function
        End Select
    Loop
    
    P_KANRI_Open = False

End Function
Public Function P_KANRI_MAKE_Proc() As Integer
'----------------------------------------------------------------------------
'                   Ç}X^Ì©®ì¬
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    P_KANRI_MAKE_Proc = True

    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)     'Úº°ÄÞ
    Call UniCode_Conv(P_KANRIREC.SHIME_DD, "31")            '÷ßú
    Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, "00000")       'w}[
    Call UniCode_Conv(P_KANRIREC.ORDER_NO, "00000")         '­
    Call UniCode_Conv(P_KANRIREC.URIAGE_NO, "00000")        'ÞãÚº°ÄÞ

    Call UniCode_Conv(P_KANRIREC.ZEI_CHANGE_YMD, "")        'ÁïÅÏXút
    Call UniCode_Conv(P_KANRIREC.NOW_ZEI_RITU, "00.0")      '»@ÁïÅ¦
    Call UniCode_Conv(P_KANRIREC.NOW_MARUME, "0")           '»@Ûß
    Call UniCode_Conv(P_KANRIREC.NEW_ZEI_RITU, "00.0")      'V@ÁïÅ¦
    Call UniCode_Conv(P_KANRIREC.NEW_MARUME, "0")           'V@Ûß

    Call UniCode_Conv(P_KANRIREC.SHONIN_CODE, "")           '³FÒº°ÄÞ
    Call UniCode_Conv(P_KANRIREC.KAISHA_NAME, "")           'ïÐ¼Ì
    Call UniCode_Conv(P_KANRIREC.TEL_NO, "")                'dbÔ
    Call UniCode_Conv(P_KANRIREC.FAX_NO, "")                'FAXÔ
    
    Call UniCode_Conv(P_KANRIREC.FILLER, "")

    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(BtOpInsert, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("¼[Åf[^gpÅ·B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "mFüÍ")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "Ç}X^")
                Exit Function
        End Select
    Loop
    
    
    P_KANRI_MAKE_Proc = False



End Function

