'Attribute VB_Name = "RefTypes"
'================================================================================================================================'
' RefTypes                                                                                                                       '
'--------------------------------------------------------                                                                        '
' https://github.com/WNKLER/RefTypes                                                                                             '
'--------------------------------------------------------                                                                        '
' A VBA/VB6 Library for reading/writing intrinsic types at arbitrary memory addresses.                                           '
' Its defining feature is that this is achieved using truly native, built-in language features.                                  '
' It uses no API declarations and has no external dependencies.                                                                  '
'================================================================================================================================'
' MIT License                                                                                                                    '
'                                                                                                                                '
' Copyright (c) 2025 Benjamin Dovidio (WNKLER)                                                                                   '
' Edited by Alexey Leonov (testuser2) 07.2025
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated                   '
' documentation files (the "Software"), to deal in the Software without restriction, including without limitation                '
' the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,                   '
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:                         '
'                                                                                                                                '
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. '
'                                                                                                                                '
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO               '
' THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE                 '
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,            '
' TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.       '
'================================================================================================================================'
Option Private Module
Option Explicit

#If Win64 = 1 Then
    Private Const Win64 As Integer = 1
    Private Const varSz As LongPtr = 24
    Private Const ptrSz As LongPtr = 8
#Else
    Private Const Win64 As Integer = 0
    Public Type LongLong
        L0x0 As Long
        L0x4 As Long
    End Type
    Private Const varSz As Long = 16
    Private Const ptrSz As Long = 4
#End If

#If VBA7 = 0 Then
    Private Enum LONG_PTR: [_]: End Enum
    Public Enum LongPtr:  [_]: End Enum '// Must be Public for Enum-typed Public Property
#End If

Private Const LPTR_SIZE As Long = 4 + (Win64 * 4)
Private Const FADF_FIXEDSIZE_AUTO = &H11&
'*********************************************************************************'
' This section just sets up the array bounds according to the four `UserDefined`  '
' parameters. If you already know the correct bounds, you can just hardcode them. '

Private Const LPROXY_SIZE      As Long = 1         '// UserDefined (should be an integer divisor of <PTR_SIZE>)
Private Const LELEMENT_SIZE    As Long = LPTR_SIZE '// UserDefined (usually should be <PTR_SIZE>)
Private Const LSIZE_PROXIED    As Long = LELEMENT_SIZE - LPROXY_SIZE
Private Const LBLOCK_STEP_SIZE As Long = LELEMENT_SIZE * LSIZE_PROXIED
Private Const Proxy_Count = 27

Private Enum BOUNDS_HELPER:
  [_BLOCK_ALLOCATION_SIZE] = Proxy_Count * LPTR_SIZE      '// UserDefined (the size of the region being proxied)
  [_0x0_ADDRESS_ALIGNMENT] = 0                     '// UserDefined (<BLOCK_ALLOCATION_ADDRESS> Mod <PTR_SIZE>)
  [_SIZE_OF_ALIGNED_BLOCK] = ([_0x0_ADDRESS_ALIGNMENT] - LPTR_SIZE) * ([_0x0_ADDRESS_ALIGNMENT] > 0) + [_BLOCK_ALLOCATION_SIZE]
  [_ELEMENT_SIZE_OF_PROXY] = ([_SIZE_OF_ALIGNED_BLOCK] + LBLOCK_STEP_SIZE - 1) \ LBLOCK_STEP_SIZE
  [_ELEMENT_SIZE_OF_BLOCK] = ([_ELEMENT_SIZE_OF_PROXY] * (LELEMENT_SIZE \ LPROXY_SIZE)) - [_ELEMENT_SIZE_OF_PROXY]
End Enum

Private Const LELEMENTS_LBOUND As Long = [_ELEMENT_SIZE_OF_PROXY] * -1
Private Const LELEMENTS_UBOUND As Long = [_ELEMENT_SIZE_OF_BLOCK] + (LELEMENTS_LBOUND < 0)
Private Const LBLOCK_OFFSET    As Long = [_ELEMENT_SIZE_OF_PROXY] * LELEMENT_SIZE '// Relative to VarPtr(<MemoryProxyVariable>)
Private Const LBLOCK_SIZE      As Long = [_ELEMENT_SIZE_OF_BLOCK] * LELEMENT_SIZE '

'*********************************************************************************'
Private Type ProxyElement               '// This could be anything. I have it as an array of bytes only for the
    ProxyAlloc(LPROXY_SIZE - 1) As Byte '// convenience of binding its size to a constant, and for maximum granularity.
End Type                                '// Although, be aware that this structure determines MemoryProxy alignment.
Private Type MemoryProxy                                           '// The declared type of `Elements()` can be any of
    Elements(LELEMENTS_LBOUND To LELEMENTS_UBOUND) As ProxyElement '// the following: Enum, UDT, or Alias (typedef)
End Type                                                           '// NOTE: A ProxyElement's Type must be smaller
Private Type B3
    b1 As Byte
    b2 As Byte
    B3 As Byte
End Type
Public Enum SortOrder
    Descending = -1
    Ascending = 1
End Enum
'// than the Type of the Element it represents.
'******************************************************************'
' When passed to `InitByProxy()`, the `Initializer.Elements` array '
' provides access to fourteen, pointer-sized elements immmediately '
' following the `Initializer` variable's memory allocation.        '
'##################################################################'
Public Initializer   As MemoryProxy
' <Memory proxied by `Initializer`>
Public iRef()     As Integer:     Private Const iRefNum = 0
Public iRef2()    As Integer:     Private Const iRef2Num = 1
Public lRef()     As Long:        Private Const lRefNum = 2
Public lRef2()    As Long:        Private Const lRef2Num = 3
Public snRef()    As Single:      Private Const snRefNum = 4
Public dRef()     As Double:      Private Const dRefNum = 5
Public cRef()     As Currency:    Private Const cRefNum = 6
Public cRef2()    As Currency:    Private Const cRef2Num = 7
Public dtRef()    As Date:        Private Const dtRefNum = 8
Public sRef()     As String:      Private Const sRefNum = 9
Public sRef2()    As String:      Private Const sRef2Num = 10
Public oRef()     As Object:      Private Const oRefNum = 11
Public blRef()    As Boolean:     Private Const blRefNum = 12
Public vRef()     As Variant:     Private Const vRefNum = 13
Public vRef2()    As Variant:     Private Const vRef2Num = 14
Public unkRef()   As IUnknown:    Private Const unkRefNum = 15
Public bRef()     As Byte:        Private Const bRefNum = 16
Public bRef2()    As Byte:        Private Const bRef2Num = 17
Public llRef()    As LongLong:    Private Const llRefNum = 18
Public lpRef()    As LongPtr:     Private Const lpRefNum = 19
Public lpRef2()   As LongPtr:     Private Const lpRef2Num = 20
Public iMap1()    As Integer:     Private Const iMap1Num = 21 'мапперы строк (с индексацией от 1)
Public iMap2()    As Integer:     Private Const iMap2Num = 22
Public bMap1()    As Byte:        Private Const bMap1Num = 23
Public bMap2()    As Byte:        Private Const bMap2Num = 24
Public b3Ref1()   As B3
Public b3Ref2()   As B3                                    '26
' <End of proxied memory block>
'##################################################################'
'******************************************************************'
'*************************************************************************************************'
' Inspired by Cristian Buse's `VBA-MemoryTools` <https://github.com/cristianbuse/VBA-MemoryTools> '
' Arbitrary memory access is achieved via a carefully constructed SAFEARRAY `Descriptor` struct.  '
Public iRef_SA As SAFEARRAY1D, _
       iRef2_SA As SAFEARRAY1D, _
       lRef_SA As SAFEARRAY1D, _
       lRef2_SA As SAFEARRAY1D, _
       snRef_SA As SAFEARRAY1D, _
       dRef_SA As SAFEARRAY1D, _
       cRef_SA As SAFEARRAY1D, _
       cRef2_SA As SAFEARRAY1D, _
       dtRef_SA As SAFEARRAY1D, _
       sRef_SA As SAFEARRAY1D, _
       sRef2_SA As SAFEARRAY1D, _
       oRef_SA As SAFEARRAY1D, _
       blRef_SA As SAFEARRAY1D, _
       vRef_SA As SAFEARRAY1D, _
       vRef2_SA As SAFEARRAY1D, _
       unkRef_SA As SAFEARRAY1D, _
       bRef_SA As SAFEARRAY1D, _
       bRef2_SA As SAFEARRAY1D, _
       llRef_SA As SAFEARRAY1D, _
       lpRef_SA As SAFEARRAY1D, _
       lpRef2_SA As SAFEARRAY1D, _
       iMap1_SA As SAFEARRAY1D, _
       iMap2_SA As SAFEARRAY1D, _
       bMap1_SA As SAFEARRAY1D, _
       bMap2_SA As SAFEARRAY1D
Public b3Ref1_SA As SAFEARRAY1D, _
       b3Ref2_SA As SAFEARRAY1D
'*************************************************************************************************'
Private IsInitialized As Boolean

Sub Initialize()
    If IsInitialized Then Exit Sub
    
    With bRef_SA
      .cDims = 1
      .fFeatures = FADF_FIXEDSIZE_AUTO
      .cLocks = 1
      .cbElements = 1
      .Bounds.cCount = 1
    End With
    bRef2_SA = bRef_SA
    iRef_SA = bRef_SA:    iRef_SA.cbElements = 2 '(LenB(iRef(0))
    iRef2_SA = bRef_SA:   iRef2_SA.cbElements = 2
    blRef_SA = bRef_SA:   blRef_SA.cbElements = 2
    lRef_SA = bRef_SA:    lRef_SA.cbElements = 4
    lRef2_SA = bRef_SA:   lRef2_SA.cbElements = 4
    cRef_SA = bRef_SA:    cRef_SA.cbElements = 8
    cRef2_SA = bRef_SA:   cRef2_SA.cbElements = 8
    snRef_SA = bRef_SA:   snRef_SA.cbElements = 4
    dRef_SA = bRef_SA:    dRef_SA.cbElements = 8
    dtRef_SA = bRef_SA:   dtRef_SA.cbElements = 8
    sRef_SA = bRef_SA:    sRef_SA.cbElements = ptrSz  ':    sRef_SA.fFeatures = 402
    sRef2_SA = bRef_SA:   sRef2_SA.cbElements = ptrSz ':    sRef2_SA.fFeatures = 402
    vRef_SA = bRef_SA:    vRef_SA.cbElements = varSz
    vRef2_SA = bRef_SA:   vRef2_SA.cbElements = varSz
    oRef_SA = bRef_SA:    oRef_SA.cbElements = ptrSz
    unkRef_SA = bRef_SA:  unkRef_SA.cbElements = ptrSz
    llRef_SA = bRef_SA:   llRef_SA.cbElements = LenB(llRef(0))
    lpRef_SA = bRef_SA:   lpRef_SA.cbElements = ptrSz
    lpRef2_SA = bRef_SA:  lpRef2_SA.cbElements = ptrSz
    iMap1_SA = bRef_SA:   iMap1_SA.cbElements = 2: iMap1_SA.Bounds.lBound = 1
    iMap2_SA = bRef_SA:   iMap2_SA.cbElements = 2: iMap2_SA.Bounds.lBound = 1
    bMap1_SA = bRef_SA:   bMap1_SA.cbElements = 1: bMap1_SA.Bounds.lBound = 1
    bMap2_SA = bRef_SA:   bMap2_SA.cbElements = 1: bMap2_SA.Bounds.lBound = 1
    b3Ref1_SA = bRef_SA:  b3Ref1_SA.cbElements = 3
    b3Ref2_SA = bRef_SA:  b3Ref2_SA.cbElements = 3
    
    InitAllByProxy Initializer.Elements, iRef_SA, Proxy_Count
    
    IsInitialized = True
End Sub
'*********************************************************************************'
' This is only possible because the compiler does not (or cannot?) discriminate   '
' between <Non-Intrinsic Array Argument> types passed to <Array Parameters> whose '
' <Declared Type> is an <Enum> or an <Alias> (a non-struct typdef).               '
' Such Array Parameters will accept any <UDT/Enum/Alias>-typed array argument.    '
'                                                                                 '
' Another key behavior is that (except for cDims, pvData, and Bounds) the array   '
' descriptor has no effect on indexing/reading/writing the array elements within  '
' the scope of the receiving procedure; indexing/reading/writing align with the   '
' declared type of the Array Parameter. (this behavior is not critical, but it    '
' greatly simplifies the implementation) NOTE: You cannot pass an element ByRef   '
' from inside the procedure. Doing so passes the address of its proxy.            '
'                                                                                 '
' Similarly, Array Parameters whose <Declared Type> is <Fixed-Length-String> will '
' accept ANY <Fixed-Length-String> array argument, regardless of Declared Length. '
' However, since Fixed-Length-Strings have no alignment, the starting position of '
' an element and the starting position of its proxy will always be the same.      '
'*********************************************************************************'
Private Sub InitAllByProxy(ProxyElements() As LONG_PTR, SA1 As SAFEARRAY1D, ByVal proxyCount&)
    Dim i&, pSA1 As LongPtr, szSA As Long
    
    pSA1 = VarPtr(SA1)
    szSA = LenB(SA1)
    For i = 0 To proxyCount - 1
        ProxyElements(i) = pSA1 + i * szSA
    Next
End Sub
Private Sub InitByProxy(ProxyElements() As LONG_PTR, ByVal num As Long, SA As SAFEARRAY1D)
    ProxyElements(num) = VarPtr(SA)
End Sub

Property Get RefInt(ByVal Target As LongPtr) As Integer
    If IsInitialized Then Else Initialize
    iRef_SA.pvData = Target
    RefInt = iRef(0)
End Property
Property Let RefInt(ByVal Target As LongPtr, ByVal RefInt As Integer)
    If IsInitialized Then Else Initialize
    iRef_SA.pvData = Target
    iRef(0) = RefInt
End Property

Property Get RefLng(ByVal Target As LongPtr) As Long
    If IsInitialized Then Else Initialize
    lRef_SA.pvData = Target
    RefLng = lRef(0)
End Property
Property Let RefLng(ByVal Target As LongPtr, ByVal RefLng As Long)
    If IsInitialized Then Else Initialize
    lRef_SA.pvData = Target
    lRef(0) = RefLng
End Property

Property Get RefSng(ByVal Target As LongPtr) As Single
    If IsInitialized Then Else Initialize
    snRef_SA.pvData = Target
    RefSng = snRef(0)
End Property
Property Let RefSng(ByVal Target As LongPtr, ByVal RefSng As Single)
    If IsInitialized Then Else Initialize
    snRef_SA.pvData = Target
    snRef(0) = RefSng
End Property

Property Get RefDbl(ByVal Target As LongPtr) As Double
    If IsInitialized Then Else Initialize
    dRef_SA.pvData = Target
    RefDbl = dRef(0)
End Property
Property Let RefDbl(ByVal Target As LongPtr, ByVal RefDbl As Double)
    If IsInitialized Then Else Initialize
    dRef_SA.pvData = Target
    dRef(0) = RefDbl
End Property

Property Get RefCur(ByVal Target As LongPtr) As Currency
    If IsInitialized Then Else Initialize
    cRef_SA.pvData = Target
    RefCur = cRef(0)
End Property
Property Let RefCur(ByVal Target As LongPtr, ByVal RefCur As Currency)
    If IsInitialized Then Else Initialize
    cRef_SA.pvData = Target
    cRef(0) = RefCur
End Property

Property Get RefDate(ByVal Target As LongPtr) As Date
    If IsInitialized Then Else Initialize
    dtRef_SA.pvData = Target
    RefDate = dtRef(0)
End Property
Property Let RefDate(ByVal Target As LongPtr, ByVal RefDate As Date)
    If IsInitialized Then Else Initialize
    dtRef_SA.pvData = Target
    dtRef(0) = RefDate
End Property

Property Get RefStr(ByVal Target As LongPtr) As String
    If IsInitialized Then Else Initialize
    sRef_SA.pvData = Target
    RefStr = sRef(0)
End Property
Property Let RefStr(ByVal Target As LongPtr, ByRef RefStr As String)
    If IsInitialized Then Else Initialize
    sRef_SA.pvData = Target
    sRef(0) = RefStr
End Property

Property Get RefObj(ByVal Target As LongPtr) As Object
    If IsInitialized Then Else Initialize
    oRef_SA.pvData = Target
    Set RefObj = oRef(0)
End Property
Property Set RefObj(ByVal Target As LongPtr, ByVal RefObj As Object)
    If IsInitialized Then Else Initialize
    oRef_SA.pvData = Target
    Set oRef(0) = RefObj
End Property

Property Get RefBool(ByVal Target As LongPtr) As Boolean
    If IsInitialized Then Else Initialize
    blRef_SA.pvData = Target
    RefBool = blRef(0)
End Property
Property Let RefBool(ByVal Target As LongPtr, ByVal RefBool As Boolean)
    If IsInitialized Then Else Initialize
    blRef_SA.pvData = Target
    blRef(0) = RefBool
End Property

Property Get RefVar(ByVal Target As LongPtr) As Variant
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    RefVar = vRef(0)
End Property
Property Let RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    vRef(0) = RefVar
End Property
Property Set RefVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    Set vRef(0) = RefVar
End Property

Property Get RefUnk(ByVal Target As LongPtr) As IUnknown
    If IsInitialized Then Else Initialize
    unkRef_SA.pvData = Target
    Set RefUnk = unkRef(0)
End Property
Property Set RefUnk(ByVal Target As LongPtr, ByVal RefUnk As IUnknown)
    If IsInitialized Then Else Initialize
    unkRef_SA.pvData = Target
    Set unkRef(0) = RefUnk
End Property

'Property Get RefDec(ByVal Target As LongPtr) As Variant
'    If IsInitialized Then Else Initialize
'    dcRef_SA.pvData = Target
'    RefDec = dcRef(0)
'End Property
'Property Let RefDec(ByVal Target As LongPtr, ByVal RefDec As Variant)
'    If IsInitialized Then Else Initialize
'    dcRef_SA.pvData = Target
'    dcRef(0) = RefDec
'End Property '_
Property Get RefByte(ByVal Target As LongPtr) As Byte
    If IsInitialized Then Else Initialize
    bRef_SA.pvData = Target
    RefByte = bRef(0)
End Property
Property Let RefByte(ByVal Target As LongPtr, ByVal RefByte As Byte)
    If IsInitialized Then Else Initialize
    bRef_SA.pvData = Target
    bRef(0) = RefByte
End Property

    Property Get RefLngLng(ByVal Target As LongPtr) As LongLong
        If IsInitialized Then Else Initialize
        llRef_SA.pvData = Target
        RefLngLng = llRef(0)
    End Property
#If Win64 = 0 Then
    Property Let RefLngLng(ByVal Target As LongPtr, ByRef RefLngLng As LongLong)
#Else
    Property Let RefLngLng(ByVal Target As LongPtr, ByVal RefLngLng As LongLong)
#End If
        If IsInitialized Then Else Initialize
        llRef_SA.pvData = Target
        llRef(0) = RefLngLng
    End Property

Property Get RefLngPtr(ByVal Target As LongPtr) As LongPtr
    If IsInitialized Then Else Initialize
    lpRef_SA.pvData = Target
    RefLngPtr = lpRef(0)
End Property
Property Let RefLngPtr(ByVal Target As LongPtr, ByVal RefLngPtr As LongPtr)
    If IsInitialized Then Else Initialize
    lpRef_SA.pvData = Target
    lpRef(0) = RefLngPtr
End Property
Property Get RefLngPtr2(ByVal Target As LongPtr) As LongPtr
    If IsInitialized Then Else Initialize
    lpRef2_SA.pvData = Target
    RefLngPtr = lpRef2(0)
End Property
Property Let RefLngPtr2(ByVal Target As LongPtr, ByVal RefLngPtr2 As LongPtr)
    If IsInitialized Then Else Initialize
    lpRef2_SA2.pvData = Target
    lpRef2(0) = RefLngPtr2
End Property
'перемещение указателя (передача владения)
Sub MovePtr(ByVal pDst As LongPtr, ByVal pSrc As LongPtr)
    If IsInitialized Then Else Initialize
    lpRef_SA.pvData = pDst
    lpRef2_SA.pvData = pSrc
    lpRef(0) = lpRef2(0)
    lpRef2(0) = 0
End Sub
Private Function StrCompVBA(str1$, str2$) As Long
    Dim len1&, len2&, lenMin&
    Dim i&, dif&
    If IsInitialized Then Else Initialize
    
    len1 = Len(str1) + 1: len2 = Len(str2) + 1
    If len1 > len2 Then lenMin = len2 Else lenMin = len1
    iRef1_SA.pvData = StrPtr(str1)
    iRef1_SA.Bounds.cCount = lenMin
    iRef2_SA.pvData = StrPtr(str2)
    iRef2_SA.Bounds.cCount = lenMin
    
    For i = 1 To lenMin
        dif = iRef1(i) - iRef2(i)
        If dif Then Exit For
    Next
    
    StrCompVBA = dif
End Function
'аналог instr$() с дополнителным параметром lStop, чтобы указывать позицию окончания поиска.
Function InStr2(sCheck$, sMatch$, Optional ByVal lStart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pvData = StrPtr(sCheck)
        iMap2_SA.pvData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        MovePtr VarPtr(sTmp1), VarPtr(StrConv(sCheck, vbLowerCase)) + 8 'v2
        MovePtr VarPtr(sTmp2), VarPtr(StrConv(sMatch, vbLowerCase)) + 8
        iMap1_SA.pvData = StrPtr(sTmp1)
        iMap2_SA.pvData = StrPtr(sTmp2)
    End If
    lenCheck = Len(sCheck)
    lenMatch = Len(sMatch)
    iMap1_SA.Bounds.cCount = lenCheck
    iMap2_SA.Bounds.cCount = lenMatch
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
            InStr2 = i: Exit Function
        End If
skip:
    Next
End Function
Function InStr2B(sCheck$, sMatch$, Optional ByVal lStart As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        bMap1_SA.pvData = StrPtr(sCheck)
        bMap2_SA.pvData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        MovePtr VarPtr(sTmp1), VarPtr(StrConv(sCheck, vbLowerCase)) + 8 'v2
        MovePtr VarPtr(sTmp2), VarPtr(StrConv(sMatch, vbLowerCase)) + 8
        bMap1_SA.pvData = StrPtr(sTmp1)
        bMap2_SA.pvData = StrPtr(sTmp2)
    End If
    lenCheck = LenB(sCheck)
    lenMatch = LenB(sMatch)
    bMap1_SA.Bounds.cCount = lenCheck
    bMap2_SA.Bounds.cCount = lenMatch
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
            InStr2B = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrRev2(sCheck$, sMatch$, Optional ByVal lStart As Long = -1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = 1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pvData = StrPtr(sCheck)
        iMap2_SA.pvData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
'        sTmp1 = StrConv(sCheck, vbLowerCase)                           'v1
'        sTmp2 = StrConv(sMatch, vbLowerCase)
        MovePtr VarPtr(sTmp1), VarPtr(StrConv(sCheck, vbLowerCase)) + 8 'v2
        MovePtr VarPtr(sTmp2), VarPtr(StrConv(sMatch, vbLowerCase)) + 8
        iMap1_SA.pvData = StrPtr(sTmp1)
        iMap2_SA.pvData = StrPtr(sTmp2)
    End If
    lenCheck = Len(sCheck)
    lenMatch = Len(sMatch)
    iMap1_SA.Bounds.cCount = lenCheck
    iMap2_SA.Bounds.cCount = lenMatch
    If lStart = -1 Then lStart = lenCheck
    
'    Dim bgnIter& '                                                     'v1
'    j = lenMatch
'    iMatch = iMap2(lenMatch)
'    bgnIter = lenMatch - 1
'    For i = lStart To lStop + lenMatch - 1 Step -1
'        If iMap1(i) <> iMatch Then
'        Else
'            k = i
'            For j = bgnIter To 1 Step -1
'                k = k - 1
'                If iMap1(k) = iMap2(j) Then Else GoTo skip
'            Next
'            InStrRev2 = k: Exit Function
'        End If
'skip:
'    Next
    iMatch = iMap2(1)                                                   'v2
    For i = lStart - lenMatch + 1 To lStop Step -1
        If iMap1(i) <> iMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If iMap1(k) = iMap2(j) Then Else GoTo skip
            Next
            InStrRev2 = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrRev2B(sCheck$, sMatch$, Optional ByVal lStart As Long = -1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal lStop As Long = 1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte
    If IsInitialized Then Else Initialize
    
    If Compare = vbBinaryCompare Then
        bMap1_SA.pvData = StrPtr(sCheck)
        bMap2_SA.pvData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
'        sTmp1 = StrConv(sCheck, vbLowerCase)                           'v1
'        sTmp2 = StrConv(sMatch, vbLowerCase)
        MovePtr VarPtr(sTmp1), VarPtr(StrConv(sCheck, vbLowerCase)) + 8 'v2
        MovePtr VarPtr(sTmp2), VarPtr(StrConv(sMatch, vbLowerCase)) + 8
        bMap1_SA.pvData = StrPtr(sTmp1)
        bMap2_SA.pvData = StrPtr(sTmp2)
    End If
    lenCheck = LenB(sCheck)
    lenMatch = LenB(sMatch)
    bMap1_SA.Bounds.cCount = lenCheck
    bMap2_SA.Bounds.cCount = lenMatch
    If lStart = -1 Then lStart = lenCheck
    
'    Dim bgnIter&                                                       'v1
'    j = lenMatch
'    bMatch = bmap2(lenMatch)
'    bgnIter = lenMatch - 1
'    For i = lStart To lStop + lenMatch - 1 Step -1
'        If bmap1(i) <> bMatch Then
'        Else
'            k = i
'            For j = bgnIter To 1 Step -1
'                k = k - 1
'                If bmap1(k) = bmap2(j) Then
'                Else
'                    GoTo skip
'                End If
'            Next
'            InStrRev2B = k: Exit Function
'        End If
'skip:
'    Next
    bMatch = bMap2(1)                                                   'v2
    For i = lStart - lenMatch + 1 To lStop Step -1
        If bMap1(i) <> bMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If bMap1(k) = bMap2(j) Then Else GoTo skip
            Next
            InStrRev2B = i: Exit Function
        End If
skip:
    Next
End Function
'Аналог CopyMemory
Sub MemLSet(ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal Size As Long)
    Dim sDst$, sSrc$, lTmp&
    Dim s1$, s2$
    If IsInitialized Then Else Initialize
    
    If Size > 3 Then
    Else
        MiniCopy pDst, pSrc, Size
        Exit Sub
    End If
    Size = Size - 4
    
    lRef_SA.pvData = pSrc
    lTmp = lRef(0)
    lRef(0) = Size
    lRef2_SA.pvData = pDst
    lRef2(0) = Size

    pSrc = pSrc + 4
    pDst = pDst + 4
    sRef_SA.pvData = VarPtr(pSrc)
    sRef2_SA.pvData = VarPtr(pDst)
    
    LSet sRef2(0) = sRef(0)
    
    lRef(0) = lTmp
    lRef2(0) = lTmp
End Sub
'вспомогательная процедура для MemLSet для копирования размера меньше 4 байт.
Sub MiniCopy(ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal Size As Long)
    On Size GoTo 1, 2, 3
    Exit Sub
    If False Then
1:
        bRef_SA.pvData = pSrc
        bRef2_SA.pvData = pDst
        bRef2(0) = bRef(0)
    ElseIf False Then
2:
        iRef_SA.pvData = pSrc
        iRef2_SA.pvData = pDst
        iRef2(0) = iRef(0)
    ElseIf False Then
3:
        b3Ref1_SA.pvData = pSrc
        b3Ref2_SA.pvData = pDst
        b3Ref2(0) = b3Ref1(0)
    End If
End Sub
Private Sub Test_ShellSortS()
    Dim sAr$()
    
    sAr = Split("яблоки Груши аппельсины Кориандр манго")
    
    ShellSortS sAr, Descending, vbTextCompare
End Sub
'http://www.excelworld.ru/board/vba/tricks/sort_array_shell/9-1-0-32
Sub ShellSortS(Arr() As String, Optional Order As SortOrder = Ascending, Optional Comp As VbCompareMethod)
    Dim Limit&, Switch&, i&, j&, ij&, Ub&
    If IsInitialized Then Else Initialize
    
    Ub = UBound(Arr)
    j = (Ub + 1) \ 2
    Do While j > 0
        Limit = Ub - j
        Do
            Switch = -1
            For i = 0 To Limit
                ij = i + j
                If StrComp(Arr(i), Arr(ij), Comp) = Order Then
                    SwapPtr VarPtr(Arr(i)), VarPtr(Arr(ij))
                    Switch = i
                End If
            Next
            Limit = Switch - j
        Loop While Switch >= 0
        j = j \ 2
    Loop
End Sub
Sub SwapPtr(ByVal p1 As LongPtr, ByVal p2 As LongPtr)
    Dim pTmp As LongPtr
    lpRef_SA.pvData = p1
    lpRef2_SA.pvData = p2
    pTmp = lpRef(0)
    lpRef(0) = lpRef2(0)
    lpRef2(0) = pTmp
End Sub
Private Sub Test_MemLSet()
    Dim s1$, s2$
    Initialize
    s1 = "1111111111"
    s2 = "2222"
'    MidB$(s1, 7, 8) = s2
'    MemLSet StrPtr(s1) + 6, StrPtr(s2), 8
'    Debug.Print s1 '1112222111
    
    Dim b1(3) As Byte, b2(3) As Byte
    b1(0) = 1: b1(1) = 2: b1(2) = 3: b1(3) = 4
    
    MemLSet VarPtr(b2(0)), VarPtr(b1(0)), 4
End Sub
Private Sub TestMovePtr()
    Dim s1$, s2$
    Dim sAr1$(), sAr2$(), lp As LongPtr
    
    Initialize
    
    s1 = "asdfadaf"
    Stop
    MovePtr VarPtr(s2), VarPtr(s1)
    Stop
    ReDim sAr1(2): sAr1(1) = "fasfsad"
    
    Stop 'см. Immediate
    MovePtr VarPtr(lp) + LenB(lp), VarPtr(lp) + LenB(lp) * 2
    Stop
End Sub
Private Sub TestStrCompVBA()
    Dim s1$, s2$, lres&, lres2&
    s1 = "abcd"
    s2 = "abc"
    lres = StrCompVBA(s2, s1)
    lres2 = StrComp(s2, s1)
End Sub
Private Sub TestInStrRev2()
    Dim sCheck$, sMatch$, lres&, lres2&, cmp As VbCompareMethod
    sCheck = "rtoiutPoIpkj"
    sMatch = "TpoI"
    cmp = TextCompare
    lres = InStrRev2(sCheck, sMatch, 9, vbTextCompare, 6)
    lres2 = InStrRev2B(sCheck, sMatch, 18, vbTextCompare, 11)
    lres = InStr2(sCheck, sMatch, 6, cmp, 9)
    lres2 = InStr2B(sCheck, sMatch, 11, cmp, 18)
    Stop
End Sub
Private Sub TestiRef()
    Dim s$: s = "АБВ"
    Initialize
    iRef_SA.pvData = StrPtr(s)
    iRef_SA.Bounds.cCount = Len(s)
    iRef(2) = AscW("Ъ")
    ReDim Preserve iRef(1 To 3)
End Sub
Private Sub TestArrayCopy()
    Dim s1$, s2$
    s1 = "АБВГД"
    s2 = "     "
    Initialize
    With iRef_SA
      .pvData = StrPtr(s1)
      .Bounds.cCount = Len(s1)
    End With
    With m_SA
      .pvData = StrPtr(s2)
      .cbElements = 2
      .Bounds.lBound = 1
      .Bounds.cCount = Len(s2)
    End With
    iRef() = iRef
    
End Sub
Private Sub TestArrayAssigment()
    Dim iAr1%(2), iAr2%(2)
'    ReDim iAr1(2)
    iAr2(0) = 123
    Debug.Print VarPtr(iAr1(0)); iAr1(0)
    LSet iAr1 = iAr2
    Debug.Print VarPtr(iAr1(0)); iAr1(0)
End Sub
'Тест проверяет устанавливает ли команда LSet нуль-терминал в конце строки. (не устанавливает)
Private Sub TestLsetString()
    Dim s1$, s2$
    Initialize
    s1 = "sffkjk"
    s2 = "      "
    iRef1_SA.pvData = StrPtr(s1)
    iRef1_SA.Bounds.cCount = Len(s1) + 1
    iRef2_SA.pvData = StrPtr(s2)
    iRef2_SA.Bounds.cCount = Len(s1) + 1
    iRef2(7) = 123
    LSet s2 = s1
    Debug.Print iRef2(7)
End Sub
Private Sub TestArray()
    Dim sAr$(2), pSA As LongPtr, SA As SAFEARRAY1D
    Initialize
    
    pSA = VarPtr(pSA) + ptrSz
'    pSA = RefLngPtr(pSA)
    CopyPtr pSA, ByVal ArrPtr(sAr)
    
    CopyMemory SA, ByVal pSA, LenB(SA)
End Sub
Private Sub Test_B3()
    Dim b3Ar(2) As B3, b3Ar2(2) As B3
    Debug.Print LenB(b3Ar(0))
    Debug.Print VarPtr(b3Ar(2)) - VarPtr(b3Ar(1))
End Sub
