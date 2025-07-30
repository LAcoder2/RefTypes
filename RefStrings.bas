Option Explicit

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
Function RefToStr(rStr%()) As String
    bMap1_SA.pData = VarPtr(rStr(1))
    bMap1_SA.Count = UBound(rStr) * 2 '.Count * 2
    RefToStr = bMap1
End Function
Private Sub Test_MidRef()
    Dim s$, rStr%(), rStrSA As SA1D
    
    s = "afpiuk44o"
    rStr = MidRef(s, 3, 4, rStrSA)
    
    Debug.Print RefToStr(rStr)
End Sub
'получение массива integer (который будет использовать на дескриптор SA), замапленного на заданную часть строки.
Function MidRef(SA As SA1D, sSrc$, ByVal start&, Optional ByVal length&) As Integer()
    Dim iArRes%(), lp As LongPtr
    If IsInitialized Then Else Initialize
    
    If LenB(sSrc) Then Else Exit Function
    
    SA = iMap1_SA
    SA.pData = StrPtr(sSrc) + (start - 1) * 2
    SA.Count = length
    PutPtr(VarPtr(lp) + ptrSz) = VarPtr(SA)
    
    MidRef = iArRes
End Function
Private Sub Test_iStrConv()
    Dim s$, rs1%(), rs1SA As SA1D, rs2%(), rs2SA As SA1D
    s = "ABCdzN"
    
    rs1 = MidRef(rs1SAs, 1, Len(s))
    
    rs2 = iStrConv(rs1)
    
    Debug.Print RefToStr(rs2)
End Sub
Function bToAnsi(iStrInp%()) As Byte()
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
    
    strLen = UBound(iStrInp)
    ReDim bStrOut(1 To strLen)
    For i = 1 To strLen
        bStrOut(i) = AnsiTbl(iStrInp(i))
    Next
    
    bToAnsi = bStrOut
End Function
Private Sub Test_toAnsi_formAnsi()
    Dim s1$, s2$, rs1%(), iUn%(), SA1 As SA1D
    Dim bAn() As Byte, b2() As Byte
    
    s1 = "asfАфЦri"
    
    rs1 = MidRef(SA1, s1, 1, Len(s1))
    bAn = bToAnsi(rs1)
    iUn = iFromAnsi(b1)
End Sub
Function iFromAnsi(bStrInp() As Byte) As Integer()
    Const ChrTblSz& = 2340
    Static init As Boolean, UnicTbl%(255)
    Dim i&, j&, strLen&, iStrOut() As Integer
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
    
    Dim Ub&, Lb&
    Lb = LBound(bStrInp)
    Ub = UBound(bStrInp)
    ReDim iStrOut(1 To Ub - Lb + 1)
    For i = Lb To Ub
        j = j + 1
        iStrOut(j) = UnicTbl(bStrInp(i))
    Next
    
    iFromAnsi = iStrOut
End Function
Function iStrConv(iStrInp%(), Optional ByVal Conv As VbStrConv) As Integer()
    Const ChrTblSz& = 2340
    Static init As Boolean, LoTbl%(0 To ChrTblSz \ 2 - 1), UpTbl%(0 To ChrTblSz \ 2 - 1)
    Dim i&, strLen&, iStrOut%()
    If init Then
    Else
        If IsInitialized Then Else Initialize
        Dim sChars$, sTmp$
        For i = 1 To 1169
            LoTbl(i) = i
        Next
        bMap1_SA.pData = VarPtr(LoTbl(0))
        bMap1_SA.Count = ChrTblSz '(1170 * 2)
        sChars = bMap1 'RefToStr(LoTbl)
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbLowerCase)) + 8
        MemLSet VarPtr(LoTbl(0)), StrPtr(sTmp), ChrTblSz
        
        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbUpperCase)) + 8
        MemLSet VarPtr(UpTbl(0)), StrPtr(sTmp), ChrTblSz
        
'        For i = 0 To 255
'            bAnsi(i) = i
'        Next
'        sChars = bAnsi
'        MovePtr VarPtr(sTmp), VarPtr(StrConv(sChars, vbUnicode)) + 8
'        MemLSet VarPtr(UnicTbl%(0)), StrPtr(sTmp), &H100 * 2
        
        init = True
    End If
    
    strLen = UBound(iStrInp)
    ReDim iStrOut(1 To strLen)
    Select Case Conv
    Case vbUpperCase
        For i = 1 To strLen
            iStrOut(i) = UpTbl(iStrInp(i))
        Next
    Case vbLowerCase
        For i = 1 To strLen
            iStrOut(i) = LoTbl(iStrInp(i))
        Next
    End Select
    
    iStrConv = iStrOut
End Function
Private Sub fsdffsfsdds()
    Dim i&, chars()
    ReDim chars(1 To 3000) ', 0 To 0)
    For i = 1 To 3000
        chars(i) = ChrW(i)
    Next
End Sub
'аналог instr$() с дополнителным параметром lStop, чтобы указывать позицию окончания поиска.
Function InStrR(rsCheck%(), rsMatch%(), Optional ByVal lstart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pData = StrPtr(sCheck)
        iMap2_SA.pData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        MovePtr VarPtr(sTmp1), VarPtr(StrConv(sCheck, vbLowerCase)) + 8 'v2
        MovePtr VarPtr(sTmp2), VarPtr(StrConv(sMatch, vbLowerCase)) + 8
        iMap1_SA.pData = StrPtr(sTmp1)
        iMap2_SA.pData = StrPtr(sTmp2)
    End If
    lenCheck = Len(sCheck)
    lenMatch = Len(sMatch)
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
            InStr2 = i: Exit Function
        End If
skip:
    Next
End Function
