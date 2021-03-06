Attribute VB_Name = "P_STOCKSUM"
Option Explicit

'********************************************************************
'*
'*              ήI΅WvΓή°ΐ  t@Cθ`
'*
'*          CREATE 2006.02.15
'********************************************************************
't@Chc
Public Const P_STOCKSUM_ID$ = "P_STOCKSUM"

'y[WTCY
Private Const P_STOCKSUM_PG_SIZ% = 1024

'|WVEubN
Public P_STOCKSUM_POS As POSBLK
'********************************************************************
'*                                                                  *
'*                           \’Μθ`                             *
'*                                                                  *
'********************************************************************
'*************************** ΪΌθ` *****************************
'R[hθ`


Public Type P_STOCKSUM_REC_Tag
    G_SYUSHI(0 To 2)            As Byte         'ϋxPΚ
    ZEN_ZAIKO_KIN(0 To 10)      As Byte         'OέΙΰz

    NYUKO_KIN(0 To 10)          As Byte         'όΙΰz
    SYUKO_KIN(0 To 10)          As Byte         'oΙΰz
    ZAIKO_KIN(0 To 10)          As Byte         '»έΙΰz
    FILLER(0 To 16)             As Byte         '


End Type
'f[^Eobt@
Public P_STOCKSUM_REC          As P_STOCKSUM_REC_Tag

'L[θ`
    
Public Type KEY0_P_STOCKSUM                    'jdxO
    G_SYUSHI(0 To 2)            As Byte         'ϋxPΚ
End Type
    
    
'L[Ef[^
Public K0_P_STOCKSUM        As KEY0_P_STOCKSUM

Type P_STOCKSUM_FSpeck
    fs                      As BtFileSpeck  ' Μ§²Ω ½Νί―Έ\’Μ
    ks0                     As BtKeySpeck   ' ·° ½Νί―Έ\’Μ
End Type

Private P_STOCKSUM_Speck       As P_STOCKSUM_FSpeck
Private Function P_STOCKSUM_Create() As Integer
'********************************************************************
'*
'*              ήI΅WvΓή°ΐ  bqd`sd
'*
'*      ψ  :Θ΅
'*      ίθl:false ³ν
'*             true  Ων
'*      ϋxΙt@CΌπͺ―ι  2007.11.13
'*
'********************************************************************
Dim sts             As Integer
Dim FullPath        As String
Dim c               As String * 128

Dim Ret             As Long     '2007.11.13




    P_STOCKSUM_Create = True
                                            'ήI΅WvΓή°ΐtpXζέ
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]ΗέέG[")
        Exit Function
    End If



    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13
   
    
    P_STOCKSUM_Speck.fs.recoleng = Len(P_STOCKSUM_REC)        ' R[h·
    P_STOCKSUM_Speck.fs.PageSize = P_STOCKSUM_PG_SIZ          ' y[WTCY
    P_STOCKSUM_Speck.fs.idexnumb = 1                       ' CfbNX
    P_STOCKSUM_Speck.fs.fileflag = 0                       ' t@CtO
    P_STOCKSUM_Speck.fs.reserve = &H0                      ' \ρΟέ
    
    '--------------------------------------------------- L[O €
    P_STOCKSUM_Speck.ks0.keypos = 1                        ' L[|WV
    P_STOCKSUM_Speck.ks0.keyleng = 3                       ' L[·
    P_STOCKSUM_Speck.ks0.keyflag = BtKfExt                 ' L[tO
    P_STOCKSUM_Speck.ks0.keytype = Chr(BtKtString)         ' L[^Cv
    P_STOCKSUM_Speck.ks0.reserve = &H0                     ' \ρΟέ
    '--------------------------------------------------- L[O ’
    
    
    sts = BTRV(BtOpCreate, P_STOCKSUM_POS, P_STOCKSUM_Speck, Len(P_STOCKSUM_Speck), ByVal FullPath, Len(FullPath), 0)
    If sts Then
        Call File_Error(sts, BtOpCreate, "ήI΅WvΓή°ΐ")
        Exit Function
    End If
    
    P_STOCKSUM_Create = False

End Function

Public Function P_STOCKSUM_Open(Mode As Integer) As Integer
'********************************************************************
'*
'*              ήI΅WvΓή°ΐ  nodm
'*
'*      ψ  :Open Mode(BtrieveQΖ)
'*      ίθl:false ³ν
'*             true  Ων
'*      ϋxΙt@CΌπͺ―ι  2007.11.13
'*
'********************************************************************
Dim sts         As Integer
Dim c           As String * 128
Dim FullPath    As String

Dim Ret             As Long     '2007.11.13


    P_STOCKSUM_Open = True
                                            'ήI΅WvΓή°ΐtpXζέ
    sts = GetIni("FILE", P_STOCKSUM_ID, "SYS", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SYS.INI [P_STOCKSUM]ΗέέG[")
        Exit Function
    End If
    '2007.11.13
'    FullPath = Trim(c)
    Ret = InStr(1, Trim(c), ".") - 1
    FullPath = Left(Trim(c), Ret) & GLB_SYUSHI_F & Right(Trim(c), Len(Trim(c)) - Ret)
    '2007.11.13

    Do
        sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
                Sleep (500&)
            Case BtErrFileNotFound
                sts = P_STOCKSUM_Create()   'ήI΅WvΓή°ΐμ¬
                If sts <> False Then
                    Exit Function
                End If
                sts = BTRV(BtOpOpen, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), ByVal FullPath, Len(FullPath), Mode)
                If sts Then
                    Call File_Error(sts, BtOpOpen, "ήI΅WvΓή°ΐ")
                    Exit Function
                End If
                Exit Do
            Case Else
                Call File_Error(sts, BtOpOpen, "ήI΅WvΓή°ΐ")
                Exit Function
        End Select
    Loop
    
    P_STOCKSUM_Open = False

End Function

