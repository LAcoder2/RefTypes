Option Explicit

#If Win64 Then
  Public Const saSz = 32
#Else
  Public Const saSz = 24
#End If

Public Type RefStr
    Buf() As Integer
    sa As SA1D
End Type

Private Sub TestSALenB()
    Dim sa As SA1D
    Debug.Print LenB(sa)
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
Function MidRef(sa As SA1D, sSrc$, Optional ByVal start& = 1, Optional ByVal length&) As Integer()
    Dim iArRes%(), lp As LongPtr, lnSrc&, maxLen&
    If IsInitialized Then Else Initialize
    
    lnSrc = Len(sSrc)
'    If lnSrc Then Else Exit Function
    If start > 0 Then Else GoTo errArgum
    If start > lnSrc Then Exit Function
    If length > 0 Then Else GoTo errArgum
    maxLen = lnSrc - start
    If length > maxLen Then length = maxLen
    
    sa = iMap1_SA
    sa.pData = StrPtr(sSrc) + (start - 1) * 2
    sa.Count = length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(sa)
    
    GoTo endFn
errArgum:
    Err.Raise 5, , "invalid function argumenct"
endFn:

    MidRef = iArRes
End Function
Private Sub Test_GetStrMap()
    Dim sAnsi$, sUnic$, istr%(), bstr() As Byte
    Dim istrSA As SA1D, bstrSA As SA1D
    
    sUnic = "лдОлЫФ"
    sAnsi = StrConv(sUnic, vbFromUnicode)
    
    istr = GetStrMap(istrSA, sUnic)
    bstr = GetStrMapB(bstrSA, sAnsi)
    
End Sub
Function GetStrMap(sa As SA1D, sInp$) As Integer()
    Dim iMap%(), lp As LongPtr, lnInp&
    If IsInitialized Then Else Initialize
    lnInp = Len(sInp)
    If lnInp Then Else Exit Function
    sa = iMap1_SA
    sa.pData = StrPtr(sInp)
    sa.Count = lnInp
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(sa)
    GetStrMap = iMap
End Function
Function GetStrMapB(sa As SA1D, sInp$) As Byte()
    Dim bMap() As Byte, lp As LongPtr, lnInp&
    If IsInitialized Then Else Initialize
    lnInp = LenB(sInp)
    If lnInp Then Else Exit Function
    sa = bMap1_SA
    sa.pData = StrPtr(sInp)
    sa.Count = lnInp
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(sa)
    GetStrMapB = bMap
End Function
Private Sub Test_IntStrConv()
    Dim s$, rs1%(), rs1SA As SA1D, rs2%(), rs2SA As SA1D
    s = "ABCdzN"
    
    rs1 = MidRef(rs1SA, s) ', 1, Len(s))
    
    rs2 = IntStrConv(rs1, vbLowerCase)
    
    Debug.Print IntToStr(rs2)
End Sub
Function IntToAnsi(IStrInp%()) As Byte()
    Const ChrTblSz& = 2340
    Static init As Boolean, AnsiTbl(0 To ChrTblSz \ 2 - 1) As Byte ', UnicTbl%(0 To 255)
    Dim i&, strLen&, bStrOut() As Byte
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
    
    strLen = UBound(IStrInp)
    ReDim bStrOut(1 To strLen)
    For i = 1 To strLen
        bStrOut(i) = AnsiTbl(IStrInp(i))
    Next
    
    IntToAnsi = bStrOut
End Function
Private Sub Test_toAnsi_formAnsi()
    Dim s1$, s2$, rs1%(), iUn%(), SA1 As SA1D
    Dim bAn() As Byte, b2() As Byte
    
    s1 = "asfАфЦri"
    
    rs1 = MidRef(SA1, s1, 1, Len(s1))
    bAn = IntToAnsi(rs1)
    iUn = IntFromAnsi(b1)
End Sub
Function IntFromAnsi(BStrInp() As Byte) As Integer()
    Const ChrTblSz& = 2340
    Static init As Boolean, UnicTbl%(255)
    Dim i&, j&, strLen&, IStrOut() As Integer
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
    lb = LBound(BStrInp)
    ub = UBound(BStrInp)
    ReDim IStrOut(1 To ub - lb + 1)
    For i = lb To ub
        j = j + 1
        IStrOut(j) = UnicTbl(BStrInp(i))
    Next
    
    IntFromAnsi = IStrOut
End Function
Private Sub TestStrConvAnsi()
    Dim sU$, sa$, sAUp$, sUUp$
    sU = "aBCd"
    sa = StrConv(sU, vbFromUnicode)
    sAUp = StrConv(sa, vbUpperCase)
    sUUp = StrConv(sAUp, vbUnicode)
    sUUp = StrConv(sa, vbUnicode)
End Sub
Private Sub Test_IntStrConv2()
    Dim s$, rs%(), sa As SA1D, isUp%()
    
    s = "abcd"
    rs = GetStrMap(sa, s)
    
    isUp = IntStrConv(rs, vbUpperCase)
    
    Debug.Print IntToStr(isUp)
End Sub
Function IntStrConv(IStrInp%(), ByVal Conv As VbStrConv) As Integer()
    Const ChrTblSz& = 2340
    Static init As Boolean, LoTbl%(0 To ChrTblSz \ 2 - 1), UpTbl%(0 To ChrTblSz \ 2 - 1)
    Dim i&, strLen&, IStrOut%()
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$
        For i = 1 To 1169
            LoTbl(i) = i
        Next
        bMap1_SA.pData = VarPtr(LoTbl(0))
        bMap1_SA.Count = ChrTblSz '(1170 * 2)
        sChars = bMap1 'IntToStr(LoTbl)
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbLowerCase)) + 8
        MemLSet VarPtr(LoTbl(0)), StrPtr(sTmp), ChrTblSz
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbUpperCase)) + 8
        MemLSet VarPtr(UpTbl(0)), StrPtr(sTmp), ChrTblSz
        
        init = True
    End If
    
    strLen = UBound(IStrInp)
    ReDim IStrOut(1 To strLen)
    Select Case Conv
    Case vbUpperCase
        For i = 1 To strLen
            IStrOut(i) = UpTbl(IStrInp(i))
        Next
    Case vbLowerCase
        For i = 1 To strLen
            IStrOut(i) = LoTbl(IStrInp(i))
        Next
    End Select
    
    IntStrConv = IStrOut
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
    Dim lres&
    
    s1 = "gdjl;eriuo":  rs1 = GetStrMap(rs1_, s1): rs3 = GetStrMapB(rs3_, s1)
    s2 = "l;er":        rs2 = GetStrMap(rs2_, s2): rs4 = GetStrMapB(rs4_, s2)
    
'    lres = InIntStr(rs1, rs2, 4, , 7)
    lres = InBytStr(rs3, rs4, 7, , 13)
    
End Sub
'аналог instr$() с дополнителным параметром lStop, чтобы указывать позицию окончания поиска.
Function InIntStr(isCheck%(), isMatch%(), Optional ByVal lstart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    If IsInitialized Then Else Initialize
    
    lpRef_SA.pData = VarPtr(lstart) - ptrSz * 2
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
    For i = lstart To lStop - lenMatch + 1
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

Function InBytStr(bsCheck() As Byte, bsMatch() As Byte, Optional ByVal lstart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte, lbCheck&, ubCheck&, lbMatch&, ubMatch&
    If IsInitialized Then Else Initialize
    
    lpRef_SA.pData = VarPtr(lstart) - ptrSz * 2
    If lpRef(0) Then Else Exit Function 'check initialization bsCheck
    lpRef_SA.pData = lpRef_SA.pData + ptrSz
    If lpRef(0) Then Else Exit Function 'check initialization bsMatch
    
    lbCheck = LBound(bsCheck): ubCheck = UBound(bsCheck)
    lbMatch = LBound(bsMatch): ubMatch = UBound(bsMatch)
    If Compare = vbBinaryCompare Then
        bMap1_SA.pData = VarPtr(bsCheck(lbCheck))
        bMap2_SA.pData = VarPtr(bsMatch(lbMatch))
    Else
        Dim bsTmp1() As Byte, bsTmp2() As Byte
        bsTmp1 = BytStrConv(bsCheck, vbUpperCase)
        bsTmp2 = BytStrConv(bsMatch, vbUpperCase)
        bMap1_SA.pData = VarPtr(bsTmp1(1))
        bMap2_SA.pData = VarPtr(bsTmp2(1))
    End If
    lenCheck = ubCheck - lbCheck
    lenMatch = ubMatch - lbMatch
    bMap1_SA.Count = lenCheck
    bMap2_SA.Count = lenMatch
    If lStop = -1 Then lStop = lenCheck
    
    bMatch = bMap2(1)                                                   'v2
    For i = lstart To lStop - lenMatch + 1
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
