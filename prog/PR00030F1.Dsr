VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00030F1 
   ClientHeight    =   10545
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   _ExtentX        =   26882
   _ExtentY        =   18600
   SectionData     =   "PR00030F1.dsx":0000
End
Attribute VB_Name = "PR00030F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      'ñæç◊ÇÃBtrieve Operation








Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "G_SYUSHI"            'CODE
    Me.Fields.Add "G_SYUSHI_N"          'ä«óùïîèê
    Me.Fields.Add "ZEN_ZAIKO_KIN"       'ëOåéç›å…ã‡äz
    Me.Fields.Add "NYUKO_KIN"           'ìñåéì¸å…ã‡äz
    Me.Fields.Add "SYUKO_KIN"           'ìñåéèoå…ã‡äz
    Me.Fields.Add "ZAIKO_KIN"           'ìñåéç›å…ã‡äz
    Me.Fields.Add "SAGAKU_KIN"          'ìñåéç∑äzã‡äz


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

    
    
    sts = BTRV(DET_com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "éëçﬁíIâµèWåv√ﬁ∞¿")
            Exit Sub
    End Select
    
    
    
    If StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode) = P_StokSum_Key Then
        Me.Fields("G_SYUSHI") = "çáÅ@Å@Å@åv"
        Me.Fields("G_SYUSHI_N") = ""
    
    Else
        Me.Fields("G_SYUSHI") = StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode)
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_STOCKSUM_REC.G_SYUSHI, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Me.Fields("G_SYUSHI_N") = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            Case BtErrKeyNotFound
                Me.Fields("G_SYUSHI_N") = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "∫∞ƒﬁœΩ¿")
                Exit Sub
        End Select
            
    End If
        
    'Clng --> Val 2016.01.08
    Me.Fields("ZEN_ZAIKO_KIN") = Format(Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)), "#,##0")
    'Clng --> Val 2016.01.08
    Me.Fields("NYUKO_KIN") = Format(Val(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)), "#,##0")
    'Clng --> Val 2016.01.08
    Me.Fields("SYUKO_KIN") = Format(Val(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)), "#,##0")
    'Clng --> Val 2016.01.08
    Me.Fields("ZAIKO_KIN") = Format(Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)), "#,##0")
    'Clng --> Val 2016.01.08
    Me.Fields("SAGAKU_KIN") = Format(Val(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode) - Val(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode))), "#,##0")
    
    

    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts As Integer

    DET_com = BtOpGetFirst
    
    
    
    
    Do
        sts = BTRV(DET_com, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
        Select Case sts
            Case BtNoErr
                If CDbl(StrConv(P_STOCKSUM_REC.ZEN_ZAIKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.NYUKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.SYUKO_KIN, vbUnicode)) = 0 And _
                    CDbl(StrConv(P_STOCKSUM_REC.ZAIKO_KIN, vbUnicode)) = 0 Then
                    sts = BTRV(BtOpDelete, P_STOCKSUM_POS, P_STOCKSUM_REC, Len(P_STOCKSUM_REC), K0_P_STOCKSUM, Len(K0_P_STOCKSUM), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpDelete, "éëçﬁíIâµèWåv√ﬁ∞¿")
                            Exit Sub
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, DET_com, "éëçﬁíIâµèWåv√ﬁ∞¿")
                Exit Sub
        End Select
        DET_com = BtOpGetNext
    Loop

    'ÉZÉìÉ^Å[ñºèÃ
    Center_Name.Caption = StrConv(P_KANRIREC.Center_Name, vbUnicode)




    DET_com = BtOpGetFirst


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    lblYY.Caption = Left(PR000301.Text1(1).Text, 4)
    lblMM.Caption = Format(Val(Right(PR000301.Text1(1).Text, 2)), "0")
    
    lblPrint_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "éëçﬁíIâµèWåvï\"

    DoEvents

End Sub

