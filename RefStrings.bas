Option Explicit

#Const SafeMode = True
#If Win64 Then
  Public Const saSz = 32
#Else
  Public Const saSz = 24
#End If

Public Type RefStr
    Buf() As Integer
    SA As SA1D
End Type

Private Sub TestSALenB()
    Dim SA As SA1D
    Debug.Print LenB(SA)
End Sub
Function IntToStr(rStr%()) As String
    bMap1_SA.pData = VarPtr(rStr(1))
    bMap1_SA.Count = UBound(rStr) * 2 '.Count * 2
    IntToStr = bMap1
End Function
Private Sub Test_MidRef()
    Dim s$, rStr%(), rStrSA As SA1D
    
    s = "afpiuk44o"
    rStr = MidRef(rStrSA, s, 3, 4)
    Debug.Print VarPtr(rStr(1))
    Debug.Print IntToStr(rStr)
End Sub
'получение массива integer (который будет использовать на дескриптор SA), замапленного на заданную часть строки.
Function MidRef(SA As SA1D, sSrc$, Optional ByVal Start& = 1, Optional ByVal Length&) As Integer()
    Dim iArRes%(), lp As LongPtr, lnSrc&
    If IsInitialized Then Else Initialize
    
    lnSrc = Len(sSrc)
  #If SafeMode Then
    Dim maxlen&
'    If lnSrc Then Else Exit Function
    If Start > 0 Then Else GoTo errArgum
    If Start > lnSrc Then Exit Function
    If Length > 0 Then Else GoTo errArgum
    maxlen = lnSrc - Start
    If Length > maxlen Then Length = maxlen
  #End If
   
    SA = iMap1_SA
    SA.pData = StrPtr(sSrc) + (Start - 1) * 2
    SA.Count = Length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(SA)
    
  #If SafeMode Then
    GoTo endFn
errArgum:
    Err.Raise 5, , "invalid function argumenct"
endFn:
  #End If
    
    MidRef = iArRes
End Function
Function MidRefB(SA As SA1D, sSrc$, ByVal Start&, Optional ByVal Length&) As Byte()
    Dim bArRes() As Byte, lp As LongPtr, lnSrc&
    If IsInitialized Then Else Initialize
    
    lnSrc = LenB(sSrc)
  #If SafeMode Then
    Dim maxlen&
'    If lnSrc Then Else Exit Function
    If Start > 0 Then Else GoTo errArgum
    If Start > lnSrc Then Exit Function
    If Length > 0 Then Else GoTo errArgum
    maxlen = lnSrc - Start
    If Length > maxlen Then Length = maxlen
  #End If
    
    SA = bMap1_SA
    SA.pData = StrPtr(sSrc) + (Start - 1) * 2
    SA.Count = Length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(SA)
    
  #If SafeMode Then
    GoTo endFn
errArgum:
    Err.Raise 5, , "invalid function argumenct"
endFn:
  #End If
  
    MidRef = bArRes
End Function
Function MidRefInt(SA As SA1D, iSrc%(), ByVal Start&, Optional ByVal Length&) As Integer()
    Dim iArRes%(), lp As LongPtr, lnSrc&
    If IsInitialized Then Else Initialize
    
    lnSrc = UBound(iSrc)
  #If SafeMode Then
    Dim maxlen&
'    If lnSrc Then Else Exit Function
    If Start > 0 Then
        If Start > lnSrc Then Exit Function
    Else: GoTo errArgum
    End If
    maxlen = lnSrc - Start + 1
    Select Case Length
    Case 0: Length = maxlen
    Case Is > 0
        If Length > maxlen Then Length = maxlen
    Case Else: GoTo errArgum 'если < 0
    End Select
  #End If
   
    SA = iMap1_SA
    SA.pData = VarPtr(iSrc(Start))
    SA.Count = Length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(SA)
    
  #If SafeMode Then
    GoTo endFn
errArgum:
    Err.Raise 5, , "invalid function argumenct"
endFn:
  #End If
    
    MidRefInt = iArRes
End Function
Function MidRefByt(SA As SA1D, bSrc%(), ByVal Start&, Optional ByVal Length&) As Byte()
    Dim bArRes() As Byte, lp As LongPtr, ubSrc&
    If IsInitialized Then Else Initialize
    
    ubSrc = UBound(bSrc)
  #If SafeMode Then
    Dim maxlen&
'    If ubSrc Then Else Exit Function
    If Start > 0 Then
        If Start > ubSrc Then Exit Function
    Else: GoTo errArgum 'if < 0
    End If
    maxlen = ubSrc - Start + 1
    Select Case Length
    Case 0: Length = maxlen
    Case Is > 0
        If Length > maxlen Then Length = maxlen
    Case Else: GoTo errArgum 'if < 0
    End Select
  #End If
   
    SA = iMap1_SA
    SA.pData = VarPtr(bSrc(Start))
    SA.Count = Length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(SA)
    
  #If SafeMode Then
    GoTo endFn
errArgum:
    Err.Raise 5, , "invalid function argumenct"
endFn:
  #End If
    
    MidRefByt = bArRes
End Function
Private Sub Test_MidRefIntByt()
    Dim i&, iAr%(0 To 8), iAr2%(), SA As SA1D
    
    For i = 0 To 8: iAr(i) = i: Next
    
    iAr2 = MidRefInt(SA, iAr, 3, 4)
End Sub
Private Sub Test_GetStrMap()
    Dim sAnsi$, sUnic$, istr%(), bstr() As Byte
    Dim istrSA As SA1D, bstrSA As SA1D
    
    sUnic = "лдОлЫФ"
    sAnsi = StrConv(sUnic, vbFromUnicode)
    
    istr = GetStrMap(istrSA, sUnic)
    bstr = GetStrMapB(bstrSA, sAnsi)
    
End Sub
Function GetStrMap(SA As SA1D, sInp$) As Integer()
    Dim iMap%(), lp As LongPtr, lnInp&
    If IsInitialized Then Else Initialize
    lnInp = Len(sInp)
'    If lnInp Then Else Exit Function
    SA = iMap1_SA
    SA.pData = StrPtr(sInp)
    SA.Count = lnInp
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(SA)
    GetStrMap = iMap
End Function
Function GetStrMapB(SA As SA1D, sInp$) As Byte()
    Dim bMap() As Byte, lp As LongPtr, lnInp&
    If IsInitialized Then Else Initialize
    lnInp = LenB(sInp)
'    If lnInp Then Else Exit Function
    SA = bMap1_SA
    SA.pData = StrPtr(sInp)
    SA.Count = lnInp
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(SA)
    GetStrMapB = bMap
End Function
Private Sub Test_IntStrConv()
    Dim s$, rs1%(), rs1SA As SA1D, rs2%(), rs2SA As SA1D
    s = "ABCdzN"
    
    rs1 = MidRef(rs1SA, s) ', 1, Len(s))
    
    rs2 = IntStrConv(rs1, vbLowerCase)
    
    Debug.Print IntToStr(rs2)
End Sub
Function IntToAnsi(IntStrInp%()) As Byte()
    Const ChrTblSz& = 2340
    Static init As Boolean, AnsiTbl(0 To ChrTblSz \ 2 - 1) As Byte ', UnicTbl%(0 To 255)
    Dim i&, strLen&, BytStrOut() As Byte
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$, iChars%(0 To ChrTblSz - 1)
        For i = 0 To ChrTblSz - 1
            iChars(i) = i
        Next
        bMap1_SA.pData = VarPtr(iChars(0))
        bMap1_SA.Count = ChrTblSz
        sChars = bMap1
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbFromUnicode)) + 8
        MemLSet VarPtr(AnsiTbl(0)), StrPtr(sTmp), ChrTblSz \ 2
        init = True
    End If
    
    strLen = UBound(IntStrInp)
    ReDim BytStrOut(1 To strLen)
    For i = 1 To strLen
        BytStrOut(i) = AnsiTbl(IntStrInp(i))
    Next
    
    IntToAnsi = BytStrOut
End Function
Private Sub Test_toAnsi_formAnsi()
    Dim s1$, s2$, rs1%(), iUn%(), SA1 As SA1D
    Dim bAn() As Byte, b2() As Byte
    
    s1 = "asfАфЦri"
    
    rs1 = MidRef(SA1, s1, 1, Len(s1))
    bAn = IntToAnsi(rs1)
    iUn = IntFromAnsi(b1)
End Sub
Function IntFromAnsi(BytStrInp() As Byte) As Integer()
    Const ChrTblSz& = 2340
    Static init As Boolean, UnicTbl%(255)
    Dim i&, j&, strLen&, IntStrOut() As Integer
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$, bAnsi(255) As Byte
        For i = 0 To 255
            bAnsi(i) = i
        Next
        bMap1_SA.pData = VarPtr(bAnsi(0))
        bMap1_SA.Count = &H100
        sChars = bMap1
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbUnicode)) + 8
        MemLSet VarPtr(UnicTbl(0)), StrPtr(sTmp), &H100 * 2
        init = True
    End If
    
    Dim ub&, lb&
    lb = LBound(BytStrInp)
    ub = UBound(BytStrInp)
    ReDim IntStrOut(1 To ub - lb + 1)
    For i = lb To ub
        j = j + 1
        IntStrOut(j) = UnicTbl(BytStrInp(i))
    Next
    
    IntFromAnsi = IntStrOut
End Function
Private Sub TestStrConvAnsi()
    Dim sU$, SA$, sAUp$, sUUp$
    sU = "aBCd"
    SA = StrConv(sU, vbFromUnicode)
    sAUp = StrConv(SA, vbUpperCase)
    sUUp = StrConv(sAUp, vbUnicode)
    sUUp = StrConv(SA, vbUnicode)
End Sub
Private Sub Test_IntStrConv2()
    Dim s$, rs%(), SA As SA1D, isUp%()
    
    s = "abcd"
    rs = GetStrMap(SA, s)
    
    isUp = IntStrConv(rs, vbUpperCase)
    
    Debug.Print IntToStr(isUp)
End Sub
Function IntStrConv(IntStrInp() As Integer, ByVal Conv As VbStrConv) As Integer()
    Const ChrTblSz& = 2340 '(1170 * 2)
    Static init As Boolean, LoTbl%(0 To ChrTblSz \ 2 - 1), UpTbl%(0 To ChrTblSz \ 2 - 1)
    Dim i&, strLen&, IntStrOut%()
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$
        For i = 1 To 1169
            LoTbl(i) = i
        Next
        bMap1_SA.pData = VarPtr(LoTbl(0))
        bMap1_SA.Count = ChrTblSz
        sChars = bMap1
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbLowerCase)) + 8
        MemLSet VarPtr(LoTbl(0)), StrPtr(sTmp), ChrTblSz
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbUpperCase)) + 8
        MemLSet VarPtr(UpTbl(0)), StrPtr(sTmp), ChrTblSz
        
        init = True
    End If
    
    strLen = UBound(IntStrInp)
    ReDim IntStrOut(1 To strLen)
    Select Case Conv
    Case vbUpperCase
        For i = 1 To strLen
            IntStrOut(i) = UpTbl(IntStrInp(i))
        Next
    Case vbLowerCase
        For i = 1 To strLen
            IntStrOut(i) = LoTbl(IntStrInp(i))
        Next
    End Select
    
    IntStrConv = IntStrOut
End Function
Private Sub Test_BytStrConv()
    Dim s$, b() As Byte, b2()
    Dim s2$
    
    s = StrConv("abcd", vbFromUnicode)
    b = s
    
    b = BytStrConv(b, vbUpperCase)
    
    s2 = StrConv(b, vbUnicode)
End Sub
Function BytStrConv(BytStrInp() As Byte, ByVal Conv As VbStrConv) As Byte()
    Const ChrTblSz& = 256
    Static init As Boolean, LoTbl(0 To ChrTblSz - 1) As Byte, UpTbl(0 To ChrTblSz - 1) As Byte
    Dim i&, BytStrOut() As Byte
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$
        For i = 1 To 255
            LoTbl(i) = i
        Next
        bMap1_SA.pData = VarPtr(LoTbl(0))
        bMap1_SA.Count = ChrTblSz
        sChars = bMap1
        sChars = StrConv(sChars, vbUnicode)
        
        sTmp = StrConv(sChars, vbLowerCase)
        sTmp = StrConv(sTmp, vbFromUnicode)
        MemLSet VarPtr(LoTbl(0)), StrPtr(sTmp), ChrTblSz
        
        sTmp = StrConv(sChars, vbUpperCase)
        sTmp = StrConv(sTmp, vbFromUnicode)
        MemLSet VarPtr(UpTbl(0)), StrPtr(sTmp), ChrTblSz
        
        init = True
    End If
    
    Dim lb&, ub&, j&
    lb = LBound(BytStrInp)
    ub = UBound(BytStrInp)
    ReDim BytStrOut(1 To ub - lb + 1)
    Select Case Conv
    Case vbUpperCase
        For i = lb To ub
            j = j + 1
            BytStrOut(j) = UpTbl(BytStrInp(i))
        Next
    Case vbLowerCase
        For i = lb To ub
            j = j + 1
            BytStrOut(j) = LoTbl(BytStrInp(i))
        Next
    End Select
    
    BytStrConv = BytStrOut
End Function
Private Sub Test_InIntStr()
    Dim s1$, s2$
    Dim rs1%(), rs2%()
    Dim rs3() As Byte, rs4() As Byte
    Dim rs1_ As SA1D, rs2_ As SA1D, rs3_ As SA1D, rs4_ As SA1D
    Dim lres1&, lres2&, lres3&, lres4&
    
    s1 = "gdjl;Eriuo":  rs1 = GetStrMap(rs1_, s1): rs3 = GetStrMapB(rs3_, s1)
    s2 = "l;er":        rs2 = GetStrMap(rs2_, s2): rs4 = GetStrMapB(rs4_, s2)
    
'    lres1 = InIntStr(rs1, rs2, 4, , 7)
'    lres2 = InBytStr(rs3, rs4, 7, , 14)
    lres3 = InIntStrRev(rs1, rs2, 7, vbTextCompare, 4)
    lres4 = InBytStrRev(rs3, rs4, 14, vbTextCompare, 7)
End Sub
'аналог InStr$() с дополнителным параметром lStop, чтобы указывать позицию окончания поиска.
Function InIntStr(isCheck%(), isMatch%(), Optional ByVal lStart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    If IsInitialized Then Else Initialize
    
    lpRef_SA.pData = VarPtr(lStart) - ptrSz * 2
    If lpRef(0) Then Else Exit Function 'check initialization isCheck
    lpRef_SA.pData = lpRef_SA.pData + ptrSz
    If lpRef(0) Then Else Exit Function 'check initialization isMatch
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pData = VarPtr(isCheck(1))
        iMap2_SA.pData = VarPtr(isMatch(1))
    Else
        Dim isTmp1%(), isTmp2%()
        isTmp1 = IntStrConv(isCheck, vbUpperCase)
        isTmp2 = IntStrConv(isMatch, vbUpperCase)
        iMap1_SA.pData = VarPtr(isTmp1(1))
        iMap2_SA.pData = VarPtr(isTmp2(1))
    End If
    
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
    lenCheck = UBound(isCheck)
    lenMatch = UBound(isMatch)
    iMap1_SA.Count = lenCheck
    iMap2_SA.Count = lenMatch
    If lStop = -1 Then lStop = lenCheck
    
    iMatch = iMap2(1)                                                   'v2
    For i = lStart To lStop - lenMatch + 1
        If iMap1(i) <> iMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If iMap1(k) = iMap2(j) Then Else GoTo skip
            Next
            InIntStr = i: Exit Function
        End If
skip:
    Next
End Function
Function InBytStr(bsCheck() As Byte, bsMatch() As Byte, Optional ByVal lStart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte, LbCheck&, UbCheck&, LbMatch&, UbMatch&
    If IsInitialized Then Else Initialize
    
    lpRef_SA.pData = VarPtr(lStart) - ptrSz * 2
    If lpRef(0) Then Else Exit Function 'check initialization bsCheck
    lpRef_SA.pData = lpRef_SA.pData + ptrSz
    If lpRef(0) Then Else Exit Function 'check initialization bsMatch
    
    LbCheck = LBound(bsCheck): UbCheck = UBound(bsCheck)
    LbMatch = LBound(bsMatch): UbMatch = UBound(bsMatch)
    If Compare = vbBinaryCompare Then
        bMap1_SA.pData = VarPtr(bsCheck(LbCheck))
        bMap2_SA.pData = VarPtr(bsMatch(LbMatch))
    Else
        Dim bsTmp1() As Byte, bsTmp2() As Byte
        bsTmp1 = BytStrConv(bsCheck, vbUpperCase)
        bsTmp2 = BytStrConv(bsMatch, vbUpperCase)
        bMap1_SA.pData = VarPtr(bsTmp1(1))
        bMap2_SA.pData = VarPtr(bsTmp2(1))
    End If
    lenCheck = UbCheck - LbCheck + 1
    lenMatch = UbMatch - LbMatch + 1
    bMap1_SA.Count = lenCheck
    bMap2_SA.Count = lenMatch
    If lStop = -1 Then lStop = lenCheck
    
    bMatch = bMap2(1)                                                   'v2
    For i = lStart To lStop - lenMatch + 1
        If bMap1(i) <> bMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If bMap1(k) = bMap2(j) Then Else GoTo skip
            Next
            InBytStr = i: Exit Function
        End If
skip:
    Next
End Function
Function InIntStrRev(isCheck%(), isMatch%(), Optional ByVal lStart As Long = -1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = 1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pData = VarPtr(isCheck(1))
        iMap2_SA.pData = VarPtr(isMatch(1))
    Else
        Dim isTmp1%(), isTmp2%()
        isTmp1 = IntStrConv(isCheck, vbUpperCase)
        isTmp2 = IntStrConv(isMatch, vbUpperCase)
        iMap1_SA.pData = VarPtr(isTmp1(1))
        iMap2_SA.pData = VarPtr(isTmp2(1))
    End If
    lenCheck = UBound(isCheck)
    lenMatch = UBound(isMatch)
    iMap1_SA.Count = lenCheck
    iMap2_SA.Count = lenMatch
    If lStart = -1 Then lStart = lenCheck
    
    iMatch = iMap2(1)                                                   'v2
    For i = lStart - lenMatch + 1 To lStop Step -1
        If iMap1(i) <> iMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If iMap1(k) = iMap2(j) Then Else GoTo skip
            Next
            InIntStrRev = i: Exit Function
        End If
skip:
    Next
End Function
Function InBytStrRev(bsCheck() As Byte, bsMatch() As Byte, Optional ByVal lStart As Long = -1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = 1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte
    Dim LbCheck&, UbCheck&, LbMatch&, UbMatch&
    If IsInitialized Then Else Initialize
    
    LbCheck = LBound(bsCheck): UbCheck = UBound(bsCheck)
    LbMatch = LBound(bsMatch): UbMatch = UBound(bsMatch)
    If Compare = vbBinaryCompare Then
        bMap1_SA.pData = VarPtr(bsCheck(LbCheck))
        bMap2_SA.pData = VarPtr(bsMatch(LbMatch))
    Else
        Dim bsTmp1() As Byte, bsTmp2() As Byte
        bsTmp1 = BytStrConv(bsCheck, vbUpperCase)
        bsTmp2 = BytStrConv(bsMatch, vbUpperCase)
        bMap1_SA.pData = VarPtr(bsTmp1(1))
        bMap2_SA.pData = VarPtr(bsTmp2(1))
    End If
    lenCheck = UbCheck - LbCheck + 1
    lenMatch = UbMatch - LbMatch + 1
    bMap1_SA.Count = lenCheck
    bMap2_SA.Count = lenMatch
    If lStart = -1 Then lStart = lenCheck
    
    bMatch = bMap2(1)                                                   'v2
    For i = lStart - lenMatch + 1 To lStop Step -1
        If bMap1(i) <> bMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If bMap1(k) = bMap2(j) Then Else GoTo skip
            Next
            InBytStrRev = i: Exit Function
        End If
skip:
    Next
End Function
