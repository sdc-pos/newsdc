VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00015F2 
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   StartUpPosition =   3  'Windows Ìùèl
   _ExtentX        =   24765
   _ExtentY        =   18600
   SectionData     =   "PI00015F2.dsx":0000
End
Attribute VB_Name = "PI00015F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Doukon_com      As Integer      '\¬^¯«ÌBtrieve Operation
Private Doukon_eof      As Integer      '\¬^¯« Eof

Private Doukon_cnt      As Integer      '\¬^¯«ÌLINE COUNT


Private SHIJI_QTY       As Double       '¡ñw¦


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "KO_NO"               'No
    Me.Fields.Add "KO_HIN_GAI"          'iÔ
    Me.Fields.Add "KO_SYUBETSU"         'íÊ
    Me.Fields.Add "KO_QTY"              'õ
    Me.Fields.Add "KO_SHIJI_QTY"        'Ê

    Me.Fields.Add "KO_ST_LOCATION"      'IÔ
    Me.Fields.Add "KO_ZAIKO_QTY"        '_ÝÉ
    Me.Fields.Add "KO_ID_NO"            'ID_NO
    Me.Fields.Add "KO_ID_BCR"           'ID_NOÊÞ°º°ÄÞ
    Me.Fields.Add "KO_BIKOU"            'õl
    Me.Fields.Add "KO_HIN_NAME"         'i¼



End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
    
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
Dim SURYO       As String

Dim ST_SOKO     As String
Dim c           As String * 128
    
    If Doukon_cnt > 19 Then
        Exit Sub
    End If
    
    Me.Fields("ko_no").Value = Doukon_Tbl_No(Doukon_cnt)
    
    If Doukon_eof Then
        Me.Fields("KO_HIN_GAI") = ""        'iÔ
        Me.Fields("KO_SYUBETSU") = ""       'íÊ
        Me.Fields("KO_QTY") = ""            'õ
        Me.Fields("KO_SHIJI_QTY") = ""      'Ê
        Me.Fields("KO_ST_LOCATION") = ""    'IÔ
        Me.Fields("KO_ZAIKO_QTY") = ""      '_ÝÉ
        Me.Fields("KO_ID_NO") = ""          'ID_NO
        Me.Fields("KO_ID_BCR") = ""         'ID_NOÊÞ°º°ÄÞ
        Me.Fields("KO_BIKOU") = ""          'õl
    
    
    Else
        sts = BTRV(Doukon_com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                    Doukon_eof = True
                End If
            Case BtErrEOF
                
                Doukon_eof = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "\¬}X^iej")
                Exit Sub
        
        End Select
                                            
        If Doukon_eof Then
            Me.Fields("KO_HIN_GAI") = ""        'iÔ
            Me.Fields("KO_SYUBETSU") = ""       'íÊ
            Me.Fields("KO_QTY") = ""            'õ
            Me.Fields("KO_SHIJI_QTY") = ""      'Ê
            Me.Fields("KO_ST_LOCATION") = ""    'IÔ
            Me.Fields("KO_ZAIKO_QTY") = ""      '_ÝÉ
            Me.Fields("KO_ID_NO") = ""          'ID_NO
            Me.Fields("KO_ID_BCR") = ""         'ID_NOÊÞ°º°ÄÞ
            Me.Fields("KO_BIKOU") = ""          'õl
            Me.Fields("KO_HIN_NAME") = ""       'i¼
                                            
                                            
                                            
        Else
                                                'iÔ
            Me.Fields("KO_HIN_GAI") = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                                                'íÊ
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "R[h}X^")
                    Exit Sub
            
            End Select
            Me.Fields("KO_SYUBETSU") = StrConv(P_CODEREC.C_RNAME, vbUnicode)
                                                'õ
            If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
            Else
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
            End If
        
        
        
            'iÚ}X^ÇÝÝ
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "iÚ}X^")
                    Exit Sub
    
            End Select
        
        
        
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Me.Fields("KO_ST_LOCATION") = ""
            Else
                'WIÔ
                
                ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                Else
                    ST_SOKO = RTrim(c)
                End If
                
                
                
                Me.Fields("KO_ST_LOCATION") = Trim(ST_SOKO) & "-" & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
        
        
            'ÝÉ
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Sub
            
            End If
            Me.Fields("KO_ZAIKO_QTY") = Format(Sumi_Qty + Mi_Qty, "#0")
            'õlORo×ÊÞ°º°h
            
            
        
            Me.Fields("KO_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
        End If
            
    
    
    
    
    
    
        Doukon_com = BtOpGetNext
    End If
    
    
    
    Doukon_cnt = Doukon_cnt + 1
    
            
    eof = False
    
    
    


End Sub

Private Sub ActiveReport_Initialize()

Dim sts             As Integer

Dim cnt             As Integer
Dim com             As Integer


Dim i               As Integer
Dim Total_Times     As Double
Dim AVE             As Double


Dim SURYO           As String

Dim ST_SOKO         As String
Dim c               As String * 128

Dim Target          As Double

Dim wkValue         As String
Dim wkEDIT_NIN      As String
Dim wkEDIT_TIMES    As String
Dim wkAVE           As String



    '\¬}X^@Í¯ÀÞ°
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")

    
    
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "\¬}X^")
            Exit Sub
    
    End Select

    'dü¯æ¼
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, Taget_SHIMUKE_CODE_KEY)
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "R[h}X^")
            Exit Sub
    
    End Select
       
    Field1.text = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))       'dü¯æ¼
    
    Field2.text = ""
    Field3.text = Format(Now, "YYYY/MM/DD HH:MM")                   '­sú

    Field4.text = ""                                                '³FÒ
    
    'SÒ
    Field5.text = ""                                                'SÒ
    
    'ûP^SÒ
    lblS_Tanto1.Visible = False
    fldS_Tanto.Visible = False
    speS_tanto1.Visible = False
    l_S_Tanto1.Visible = False
    fldS_Tanto.text = ""                                        'ûP^SÒ
    
    
    lblSHIJI_F.Caption = ""
    
    
    
    
    
    
    Field7.text = Taget_Hin_key                                 'iÔ
                                                                'Ê
    SHIJI_QTY = SHIJI_QTY
    Field8.text = Format(SHIJI_QTY, "#0")
    'i¼^IÔ
    Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "iÚ}X^")
            Exit Sub
    
    End Select
    Field9.text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                      'i¼

    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        Field10.text = ""                                                   'WIÔ
    Else
        ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
        Else
            ST_SOKO = RTrim(c)
        End If
        
        
        Field10.text = Trim(ST_SOKO) & "-" & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode)
    End If

    Field11.text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))       '¤i»¸×½
    Field12.text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))    'tÁ¸×½
    Field13.text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))    'àE¸×½


    'x\tvãL³
    If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_OFF Then
        lblLabel_NIN.Caption = "******"
        lblLabel_TIMES.Caption = "******"
    Else
        lblLabel_NIN.Caption = ""
        lblLabel_TIMES.Caption = ""
    End If


    Field14.text = ""                                                       '¤i»èzæ
    

    'ÂÞÌ[v
    cnt = 0

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "\¬}X^iqj")
                Exit Sub
        
        End Select
        'iÚ}X^ÇÝÝ
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "iÚ}X^")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                'ÂÞ
                Field15.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'ÂÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field16.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field16.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
        
                Field17.text = ""
                
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field18.text = ""
                Else
                    
                    
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    
                    Field18.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                'ÂÞ
                Field19.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'ÂÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field20.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field20.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field21.text = ""
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field22.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field22.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                'ÂÞ
                Field23.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'ÂÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field24.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field24.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field25.text = ""
                
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field26.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field26.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
            Case 4
            
                'ÂÞ
                Field27.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'ÂÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field28.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field28.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field29.text = ""
                
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field30.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field30.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 5
                'ÂÞ
                Field31.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'ÂÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field32.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field32.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                Field33.text = ""
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field34.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field34.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
        
        End Select
        com = BtOpGetNext
    
    Loop


    'OÞÌ[v
    cnt = 0

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "\¬}X^iqj")
                Exit Sub
        
        End Select
        'iÚ}X^ÇÝÝ
        Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
        Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "iÚ}X^")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                'OÞ
                Field35.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'OÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field36.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field36.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                
                
                Field37.text = ""
                
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field38.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field38.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                'OÞ
                Field39.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'OÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field40.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field40.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field41.text = ""
                
                
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field42.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field42.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                'OÞ
                Field43.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                'OÞ@õ
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field44.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field44.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                Field45.text = ""
                
                'WIÔ
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field46.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field46.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
        
        End Select
    
        com = BtOpGetNext
    
    Loop

    Field47.text = Trim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))               'õl
    
    
        
        
    
    
    '©{ì¬ÌóL³
    lblSample.Visible = False
    Shape10.Visible = False

    
    
    'Ò²ÝÊÞ°º°ÄÞ
    fldMain_Bcr.text = ""

    
    '¾×õl
    fldBIKOU.Visible = False

    
    fldSyuka_No.Visible = False
    fldSyuka_Bcr.Visible = False

    '¯«i
    lblDOUKON.Visible = PRI_DOUKON
    lblDOUKON_GOUHI.Visible = PRI_DOUKON



    '\¬^¯«
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_NO, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    'üÉ®¹ó
    l_Nyuko_IN1.Visible = PRI_NYUKO_IN
    l_Nyuko_IN2.Visible = PRI_NYUKO_IN
    l_Nyuko_IN3.Visible = PRI_NYUKO_IN
    l_Nyuko_IN4.Visible = PRI_NYUKO_IN

    lblNyuko_In.Visible = PRI_NYUKO_IN

    'üÍ®¹ó
    l_Input_IN1.Visible = PRI_INPUT_IN
    l_Input_IN2.Visible = PRI_INPUT_IN
    l_Input_IN3.Visible = PRI_INPUT_IN
    l_Input_IN4.Visible = PRI_INPUT_IN

    lblInput_In.Visible = PRI_INPUT_IN


    If Not PRI_NYUKO_IN And Not PRI_NYUKO_IN Then
        l_IN_Center.Visible = False
    Else
        l_IN_Center.Visible = True
    End If

    lblBunnou.Visible = False

    

    Field60.Visible = PRI_HINBAN_BIKOU
    
    Field60.text = StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)           'iÔ



    '©Ó^Cg
    lblBEF_JISSEKI.Caption = ""
    lblBEF_BEFORE1.Caption = ""
    lblBEF_BEFORE2.Caption = ""
    lblBEF_BEFORE3.Caption = ""
    lblBEF_BEFORE4.Caption = ""
    lblBEF_SAGYO1.Caption = ""
    lblBEF_SAGYO2.Caption = ""
    lblBEF_SAGYO3.Caption = ""
    lblBEF_AFTER1.Caption = ""
    lblBEF_AFTER2.Caption = ""
    lblBEF_JISEKI.Caption = ""
    lblBEF_TASEKI.Caption = ""
            

    lblBunnou.Visible = False


    lblTarget1.Visible = True
    lblTarget2.Visible = True
    

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")


    Doukon_com = BtOpGetGreater
    Doukon_eof = False

    Doukon_cnt = 0



End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 10
    Me.PageTopMargin = 10
    Me.PageLeftMargin = 20
    Me.PageRightMargin = 20

    Me.documentName = "¤i»w}[F"

    DoEvents

End Sub

