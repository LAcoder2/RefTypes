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
Private Const FADF_STATIC As Long = &H2
Private Const FADF_FIXEDSIZE As Long = &H10
Private Const FADF_FIXEDSIZE_AUTO = &H11&
'*********************************************************************************'
' This section just sets up the array bounds according to the four `UserDefined`  '
' parameters. If you already know the correct bounds, you can just hardcode them. '

Private Const LPROXY_SIZE      As Long = 1         '// UserDefined (should be an integer divisor of <PTR_SIZE>)
Private Const LELEMENT_SIZE    As Long = LPTR_SIZE '// UserDefined (usually should be <PTR_SIZE>)
Private Const LSIZE_PROXIED    As Long = LELEMENT_SIZE - LPROXY_SIZE
Private Const LBLOCK_STEP_SIZE As Long = LELEMENT_SIZE * LSIZE_PROXIED
Private Const Proxy_Count = 1 '28 '!!!!!!!!!!!!!!!

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
'Public Type SAFEARRAYBOUND
'    cCount              As Long
'    lBound              As Long
'End Type
'Public Type SAFEARRAY1D
'    cDims               As Integer
'    fFeatures           As Integer
'    cbElements          As Long
'    cLocks              As Long
'  #If Win64 Then
'    padding             As Long
'  #End If
'    pvData              As LongPtr
'    Bounds              As SAFEARRAYBOUND
'End Type
Private Type B3
    b1 As Byte
    b2 As Byte
    B3 As Byte
End Type
Private Type lpVariant
    vt As Integer
    iunuse As Integer
    lunuse As Long
    val As LongPtr
    lpunuse As LongPtr
End Type
Private Type sVariant
    vt As Integer
    iunuse As Integer
    lunuse As Long
    val As String
    lpunuse As LongPtr
End Type
Private Type dVariant
    vt As Integer
    iunuse As Integer
    lunuse As Long
    val As Double
  #If Win64 Then
    lpunuse As LongPtr
  #End If
End Type
Private Type dtVariant
    vt As Integer
    iunuse As Integer
    lunuse As Long
    val As Date
  #If Win64 Then
    lpunuse As LongPtr
  #End If
End Type
Private Type iVariant
    vt As Integer
    iunuse As Integer
    lunuse As Long
    val As Integer
  #If Win64 Then
    lppadding As LongPtr
  #End If
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
Public lpRef() As LongPtr, lpRef_SA As SAFEARRAY1D
Public lpRef2() As LongPtr, lpRef2_SA As SAFEARRAY1D
Public iRef() As Integer, iRef_SA As SAFEARRAY1D
Public iRef2() As Integer, iRef2_SA As SAFEARRAY1D
Public lRef() As Long, lRef_SA As SAFEARRAY1D
Public lRef2() As Long, lRef2_SA As SAFEARRAY1D
Public snRef() As Single, snRef_SA As SAFEARRAY1D
Public dRef() As Double, dRef_SA As SAFEARRAY1D
Public cRef() As Currency, cRef_SA As SAFEARRAY1D
Public cRef2() As Currency, cRef2_SA As SAFEARRAY1D
Public dtRef() As Date, dtRef_SA As SAFEARRAY1D
Public sRef() As String, sRef_SA As SAFEARRAY1D
Public sRef2() As String, sRef2_SA As SAFEARRAY1D
Public oRef() As Object, oRef_SA As SAFEARRAY1D
Public blRef() As Boolean, blRef_SA As SAFEARRAY1D

Public vRef() As Variant, vRef_SA As SAFEARRAY1D
Public vlpRef() As lpVariant
Public vdRef() As dVariant
Public vdtRef() As dtVariant
Public viRef() As iVariant
'Public vsRef() As sVariant '.cLocks = 1 does not work

Public vRef2() As Variant, vRef2_SA As SAFEARRAY1D
Public unkRef() As IUnknown, unkRef_SA As SAFEARRAY1D
Public bRef() As Byte, bRef_SA As SAFEARRAY1D
Public bRef2() As Byte, bRef2_SA As SAFEARRAY1D
Public llRef() As LongLong, llRef_SA As SAFEARRAY1D
Public iMap1() As Integer, iMap1_SA As SAFEARRAY1D      'мапперы строк (с индексацией от 1)
Public iMap2() As Integer, iMap2_SA As SAFEARRAY1D
Public bMap1() As Byte, bMap1_SA As SAFEARRAY1D
Public bMap2() As Byte, bMap2_SA As SAFEARRAY1D
Public b3Ref1() As B3, b3Ref1_SA As SAFEARRAY1D
Public b3Ref2() As B3, b3Ref2_SA As SAFEARRAY1D         '26
Public saRef() As SAFEARRAY1D, saRef_SA As SAFEARRAY1D  '27


' <End of proxied memory block>
'##################################################################'
'******************************************************************'
'*************************************************************************************************'
' Inspired by Cristian Buse's `VBA-MemoryTools` <https://github.com/cristianbuse/VBA-MemoryTools> '
' Arbitrary memory access is achieved via a carefully constructed SAFEARRAY `Descriptor` struct.  '
'*************************************************************************************************'
Private IsInitialized As Boolean, islpRefInit As Boolean
Private iMapDyn_SA As SAFEARRAY1D, bMapDyn_SA As SAFEARRAY1D

Sub Initialize()
    Dim pArr As LongPtr
    If IsInitialized Then Exit Sub
    
    With lpRef_SA
      .cDims = 1
      .fFeatures = FADF_FIXEDSIZE_AUTO
      .cLocks = 1
      .cbElements = ptrSz
      .Bounds.cCount = 1
    End With
    InitByProxy Initializer.Elements, lpRef_SA, 1 'Proxy_Count инициализация первой ссылки (lpRef())
    islpRefInit = True
    
    MakeRef lpRef2_SA, VarPtr(lpRef2_SA) - ptrSz, ptrSz
'    lpRef2 = RefPtr(lpRef2_SA)
    MakeRef iRef_SA, VarPtr(iRef_SA) - ptrSz, 2
    MakeRef iRef2_SA, VarPtr(iRef2_SA) - ptrSz, 2
    MakeRef blRef_SA, VarPtr(blRef_SA) - ptrSz, 2
    MakeRef lRef_SA, VarPtr(lRef_SA) - ptrSz, 4
    MakeRef lRef2_SA, VarPtr(lRef2_SA) - ptrSz, 4
    MakeRef cRef_SA, VarPtr(cRef_SA) - ptrSz, 8
    MakeRef cRef2_SA, VarPtr(cRef2_SA) - ptrSz, 8
    MakeRef snRef_SA, VarPtr(snRef_SA) - ptrSz, 4
    MakeRef dRef_SA, VarPtr(dRef_SA) - ptrSz, 8
    MakeRef dtRef_SA, VarPtr(dtRef_SA) - ptrSz, 8
    MakeRef sRef_SA, VarPtr(sRef_SA) - ptrSz, ptrSz
    MakeRef sRef2_SA, VarPtr(sRef2_SA) - ptrSz, ptrSz
    
    MakeRef vRef_SA, VarPtr(vRef_SA) - ptrSz, varSz 'vRef()
    pArr = VarPtr(vRef_SA) + LenB(vRef_SA)
      PutPtr(pArr) = VarPtr(vRef_SA) 'vlpRef()
    pArr = pArr + ptrSz
      PutPtr(pArr) = VarPtr(vRef_SA) 'vdRef()
    pArr = pArr + ptrSz
      PutPtr(pArr) = VarPtr(vRef_SA) 'vdtRef()
    pArr = pArr + ptrSz
      PutPtr(pArr) = VarPtr(vRef_SA) 'viRef()
    
    MakeRef vRef2_SA, VarPtr(vRef2_SA) - ptrSz, varSz
    MakeRef oRef_SA, VarPtr(oRef_SA) - ptrSz, ptrSz
    MakeRef unkRef_SA, VarPtr(unkRef_SA) - ptrSz, ptrSz
    MakeRef llRef_SA, VarPtr(llRef_SA) - ptrSz, 8
    MakeRef iMap1_SA, VarPtr(iMap1_SA) - ptrSz, 2: iMap1_SA.Bounds.lBound = 1 'мапперы строк
    MakeRef iMap2_SA, VarPtr(iMap2_SA) - ptrSz, 2: iMap2_SA.Bounds.lBound = 1
    MakeRef bMap1_SA, VarPtr(bMap1_SA) - ptrSz, 1: bMap1_SA.Bounds.lBound = 1
    MakeRef bMap2_SA, VarPtr(bMap2_SA) - ptrSz, 1: bMap2_SA.Bounds.lBound = 1
    MakeRef b3Ref1_SA, VarPtr(b3Ref1_SA) - ptrSz, 3                           'ссылка 3-байтного типа
    MakeRef b3Ref2_SA, VarPtr(b3Ref2_SA) - ptrSz, 3
    MakeRef saRef_SA, VarPtr(saRef_SA) - ptrSz, LenB(saRef_SA) 'ссылка на структуру SafeArray
    
    iMapDyn_SA = iRef_SA: iMapDyn_SA.cLocks = 0: iMapDyn_SA.fFeatures = 128
    bMapDyn_SA = bRef_SA: bMapDyn_SA.cLocks = 0: bMapDyn_SA.fFeatures = 128
    
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
Private Sub InitByProxy(ProxyElements() As LONG_PTR, SA1 As SAFEARRAY1D, ByVal proxyCount&)
    Dim i&, pSA1 As LongPtr, szSA As Long
    
    pSA1 = VarPtr(SA1)
    szSA = LenB(SA1)
    For i = 0 To proxyCount - 1
        ProxyElements(i) = pSA1 + i * szSA
    Next
End Sub
'Private Sub InitByProxy(ProxyElements() As LONG_PTR, ByVal num As Long, SA As SAFEARRAY1D)
'    ProxyElements(num) = VarPtr(SA)
'End Sub

'>>>>>>>>>>>>>>MEMORY SECTION<<<<<<<<<<<<<<<<
Sub MakeRef(SA As SAFEARRAY1D, ByVal pArrOut As LongPtr, ByVal ElemSize As LongPtr)
'    Dim pTmp As LongPtr
    If pArrOut > 0 Then Else Exit Sub
    If islpRefInit Then Else Initialize
    
'    pTmp = lpRef_SA.pvData
    SA = lpRef_SA
    SA.cbElements = ElemSize
    lpRef_SA.pvData = pArrOut
    lpRef(0) = VarPtr(SA)
'    lpRef_SA.pvData = pTmp
End Sub

Property Get GetInt(ByVal Target As LongPtr) As Integer
    If IsInitialized Then Else Initialize
    iRef_SA.pvData = Target
    RefInt = iRef(0)
End Property
Property Let PutInt(ByVal Target As LongPtr, ByVal RefInt As Integer)
    If IsInitialized Then Else Initialize
    iRef_SA.pvData = Target
    iRef(0) = RefInt
End Property

Property Get GetLng(ByVal Target As LongPtr) As Long
    If IsInitialized Then Else Initialize
    lRef_SA.pvData = Target
    RefLng = lRef(0)
End Property
Property Let PutLng(ByVal Target As LongPtr, ByVal RefLng As Long)
    If IsInitialized Then Else Initialize
    lRef_SA.pvData = Target
    lRef(0) = RefLng
End Property

Property Get GetSng(ByVal Target As LongPtr) As Single
    If IsInitialized Then Else Initialize
    snRef_SA.pvData = Target
    RefSng = snRef(0)
End Property
Property Let PutSng(ByVal Target As LongPtr, ByVal RefSng As Single)
    If IsInitialized Then Else Initialize
    snRef_SA.pvData = Target
    snRef(0) = RefSng
End Property

Property Get GetDbl(ByVal Target As LongPtr) As Double
    If IsInitialized Then Else Initialize
    dRef_SA.pvData = Target
    RefDbl = dRef(0)
End Property
Property Let PutDbl(ByVal Target As LongPtr, ByVal RefDbl As Double)
    If IsInitialized Then Else Initialize
    dRef_SA.pvData = Target
    dRef(0) = RefDbl
End Property

Property Get GetCur(ByVal Target As LongPtr) As Currency
    If IsInitialized Then Else Initialize
    cRef_SA.pvData = Target
    RefCur = cRef(0)
End Property
Property Let PutCur(ByVal Target As LongPtr, ByVal RefCur As Currency)
    If IsInitialized Then Else Initialize
    cRef_SA.pvData = Target
    cRef(0) = RefCur
End Property

Property Get GetDate(ByVal Target As LongPtr) As Date
    If IsInitialized Then Else Initialize
    dtRef_SA.pvData = Target
    RefDate = dtRef(0)
End Property
Property Let PutDate(ByVal Target As LongPtr, ByVal RefDate As Date)
    If IsInitialized Then Else Initialize
    dtRef_SA.pvData = Target
    dtRef(0) = RefDate
End Property

Property Get GetStr(ByVal Target As LongPtr) As String
    If IsInitialized Then Else Initialize
    sRef_SA.pvData = Target
    GetStr = sRef(0)
End Property
Property Let PutStr(ByVal Target As LongPtr, ByRef PutStr As String)
    If IsInitialized Then Else Initialize
    sRef_SA.pvData = Target
    sRef(0) = PutStr
End Property
Function RefStr(SA As SAFEARRAY1D, Optional ByVal pData As LongPtr) As String()
    Dim lpArTmp() As String, pTmp As LongPtr
    If islpRefInit Then Else Initialize
    
'    pTmp = lpRef_SA.pvData
    SA = sRef_SA
    If pData > 0 Then SA.pvData = pData
    lpRef_SA.pvData = VarPtr(pTmp) + ptrSz
    lpRef(0) = VarPtr(SA)
'    lpRef_SA.pvData = pTmp
    
    RefStr = lpArTmp
End Function

Property Get GetObj(ByVal Target As LongPtr) As Object
    If IsInitialized Then Else Initialize
    oRef_SA.pvData = Target
    Set RefObj = oRef(0)
End Property
Property Set SetObj(ByVal Target As LongPtr, ByVal RefObj As Object)
    If IsInitialized Then Else Initialize
    oRef_SA.pvData = Target
    Set oRef(0) = RefObj
End Property

Property Get GetBool(ByVal Target As LongPtr) As Boolean
    If IsInitialized Then Else Initialize
    blRef_SA.pvData = Target
    RefBool = blRef(0)
End Property
Property Let PutBool(ByVal Target As LongPtr, ByVal RefBool As Boolean)
    If IsInitialized Then Else Initialize
    blRef_SA.pvData = Target
    blRef(0) = RefBool
End Property

Property Get GetVar(ByVal Target As LongPtr) As Variant
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    RefVar = vRef(0)
End Property
Property Let PutVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    vRef(0) = RefVar
End Property
Property Set SetVar(ByVal Target As LongPtr, ByRef RefVar As Variant)
    If IsInitialized Then Else Initialize
    vRef_SA.pvData = Target
    Set vRef(0) = RefVar
End Property

Property Get GetUnk(ByVal Target As LongPtr) As IUnknown
    If IsInitialized Then Else Initialize
    unkRef_SA.pvData = Target
    Set RefUnk = unkRef(0)
End Property
Property Set SetUnk(ByVal Target As LongPtr, ByVal RefUnk As IUnknown)
    If IsInitialized Then Else Initialize
    unkRef_SA.pvData = Target
    Set unkRef(0) = RefUnk
End Property

'Property Get GetDec(ByVal Target As LongPtr) As Variant
'    If IsInitialized Then Else Initialize
'    dcRef_SA.pvData = Target
'    RefDec = dcRef(0)
'End Property
'Property Let PutDec(ByVal Target As LongPtr, ByVal RefDec As Variant)
'    If IsInitialized Then Else Initialize
'    dcRef_SA.pvData = Target
'    dcRef(0) = RefDec
'End Property '_
Property Get GetByte(ByVal Target As LongPtr) As Byte
    If IsInitialized Then Else Initialize
    bRef_SA.pvData = Target
    RefByte = bRef(0)
End Property
Property Let PutByte(ByVal Target As LongPtr, ByVal RefByte As Byte)
    If IsInitialized Then Else Initialize
    bRef_SA.pvData = Target
    bRef(0) = RefByte
End Property

    Property Get GetLngLng(ByVal Target As LongPtr) As LongLong
        If IsInitialized Then Else Initialize
        llRef_SA.pvData = Target
        RefLngLng = llRef(0)
    End Property
#If Win64 = 0 Then
    Property Let PutLngLng(ByVal Target As LongPtr, ByRef RefLngLng As LongLong)
#Else
    Property Let PutLngLng(ByVal Target As LongPtr, ByVal RefLngLng As LongLong)
#End If
        If IsInitialized Then Else Initialize
        llRef_SA.pvData = Target
        llRef(0) = RefLngLng
    End Property

Property Get GetPtr(ByVal Target As LongPtr) As LongPtr
    If islpRefInit Then Else Initialize
    lpRef_SA.pvData = Target
    GetPtr = lpRef(0)
End Property
Property Let PutPtr(ByVal Target As LongPtr, ByVal PutPtr As LongPtr)
    If islpRefInit Then Else Initialize
    lpRef_SA.pvData = Target
    lpRef(0) = PutPtr
End Property
Function RefPtr(SA As SAFEARRAY1D, Optional ByVal Target As LongPtr) As LongPtr()
    Dim lpArTmp() As LongPtr, pTmp As LongPtr
    If islpRefInit Then Else Initialize
    
'    pTmp = lpRef_SA.pvData
    
    SA = lpRef_SA
    SA.pvData = Target
    lpRef_SA.pvData = VarPtr(pTmp) + ptrSz
    lpRef(0) = VarPtr(SA)
    
'    lpRef_SA.pvData = pTmp
    
    RefPtr = lpArTmp
End Function

Property Get GetSA(ByVal Target As LongPtr) As SAFEARRAY1D
    If IsInitialized Then Else Initialize
    saRef_SA.pvData = Target
    GetSA = saRef(0)
End Property
Property Let PutSA(ByVal Target As LongPtr, RefSA As SAFEARRAY1D)
    If IsInitialized Then Else Initialize
    saRef_SA.pvData = Target
    saRef(0) = PutSA
End Property

'перемещение указателя (передача владения)
Sub MovePtr(ByVal pDst As LongPtr, ByVal pSrc As LongPtr)
    If IsInitialized Then Else Initialize
    lpRef_SA.pvData = pDst
    lpRef2_SA.pvData = pSrc
    lpRef(0) = lpRef2(0)
    lpRef2(0) = 0
End Sub
'обмен указателями
Sub SwapPtr(ByVal p1 As LongPtr, ByVal p2 As LongPtr)
    Dim pTmp As LongPtr
    lpRef_SA.pvData = p1
    lpRef2_SA.pvData = p2
    pTmp = lpRef(0)
    lpRef(0) = lpRef2(0)
    lpRef2(0) = pTmp
End Sub

'Аналог CopyMemory
Sub MemLSet(ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal size As Long)
    Dim sDst$, sSrc$, lTmp&
    Dim s1$, s2$
    If IsInitialized Then Else Initialize
    
    If size > 3 Then
    Else
        MiniCopy pDst, pSrc, size
        Exit Sub
    End If
    size = size - 4
    
    lRef_SA.pvData = pSrc
    lTmp = lRef(0)
    lRef(0) = size
    lRef2_SA.pvData = pDst
    lRef2(0) = size

    pSrc = pSrc + 4
    pDst = pDst + 4
    sRef_SA.pvData = VarPtr(pSrc)
    sRef2_SA.pvData = VarPtr(pDst)
    
    LSet sRef2(0) = sRef(0)
    
    lRef(0) = lTmp
    lRef2(0) = lTmp
End Sub
'вспомогательная процедура для MemLSet для копирования размера меньше 4 байт.
Sub MiniCopy(ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal size As Long)
    On size GoTo 1, 2, 3
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
Function VbaMemRealloc(ByVal pBgn As LongPtr, ByVal newSize As Long) As LongPtr
    Dim bMap() As Byte, lp As LongPtr
    If newSize < 1 Then Exit Function
    If IsInitialized Then Else Initialize
    
    If pBgn Then
    Else
        ReDim bMap(newSize - 1)
        lpRef_SA.pvData = VarPtr(lp) + ptrSz
        saRef_SA.pvData = lpRef(0)
        VbaMemRealloc = saRef(0).pvData ' = VarPtr(bMap(0))
        saRef(0).pvData = 0
        Exit Function
    End If
    
    bMapDyn_SA.pvData = pBgn
    bMapDyn_SA.Bounds.cCount = newSize
    lpRef_SA.pvData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(bMapDyn_SA)
    ReDim Preserve bMap(newSize - 1)
    lpRef(0) = 0
    VbaMemRealloc = bMapDyn_SA.pvData
End Function
Function VbaMemAlloc(ByVal size As LongPtr) As LongPtr
    Dim bMap() As Byte, lp As LongPtr
    If IsInitialized Then Else Initialize
    
    ReDim bMap(size - 1)
    lpRef_SA.pvData = VarPtr(lp) + ptrSz
    saRef_SA.pvData = lpRef(0)
    With saRef(0)
      VbaMemAlloc = .pvData ' = VarPtr(bMap(0))
      .pvData = 0
    End With
End Function
'Sub VbaMemFree2(ByVal ptr As LongPtr)
'    Dim bMap() As Byte, lp As LongPtr
'    bMap = vbNullString 'Array()
'    lpRef_SA.pvData = VarPtr(lp) + ptrSz
'    saRef_SA.pvData = lpRef(0)
'    saRef(0).pvData = ptr
'End Sub
Sub VbaMemFree(ByVal ptr As LongPtr)
    Dim s$
    lpRef_SA.pvData = VarPtr(s)
    lpRef(0) = ptr + 4
End Sub

'>>>>>>>>>>>>>>>STRINGS SECTION<<<<<<<<<<<<<<<<<<<'
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
Private Sub TestResizeString()
    Dim s$: s = "abcd"
    
    ReallocStringB s, 11
End Sub
Sub ReallocString(sSrc$, ByVal newSize&)
    Dim iMap%(), pSrc As LongPtr
    If newSize < 0 Then Exit Sub
    If IsInitialized Then Else Initialize
        
    pSrc = StrPtr(sSrc)
    If pSrc Then
        iMapDyn_SA.pvData = StrPtr(sSrc) - 4
        iMapDyn_SA.Bounds.cCount = Len(sSrc) + 3
    Else
        sSrc = String$(newSize, vbNullChar)
        Exit Sub
    End If
        
    lpRef_SA.pvData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(iMapDyn_SA)
    
    ReDim Preserve iMap(newSize + 2)
    lpRef(0) = 0
    
    lpRef2_SA.pvData = VarPtr(sSrc)
    lpRef2(0) = iMapDyn_SA.pvData + 4
    
    lRef_SA.pvData = iMapDyn_SA.pvData
    lRef(0) = newSize * 2
End Sub
Sub ReallocStringB(sSrc$, ByVal newSize&)
    Dim bMap() As Byte, pSrc As LongPtr
    If newSize < 0 Then Exit Sub
    If IsInitialized Then Else Initialize
        
    pSrc = StrPtr(sSrc)
    If pSrc Then
    Else
        ReDim bMap(newSize - 1) As Byte
        sSrc = bMap
        Exit Sub
    End If
    bMapDyn_SA.pvData = StrPtr(sSrc) - 4
    bMapDyn_SA.Bounds.cCount = LenB(sSrc) + 6
        
    lpRef_SA.pvData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(bMapDyn_SA)
    
    ReDim Preserve bMap(newSize + 5)
    lpRef(0) = 0
    
    lpRef2_SA.pvData = VarPtr(sSrc)
    lpRef2(0) = bMapDyn_SA.pvData + 4
    
    lRef_SA.pvData = bMapDyn_SA.pvData
    lRef(0) = newSize
End Sub
'аналог SysAllocStringLen
Function VbaMemAllocStringLen(ByVal pStr As LongPtr, ByVal strLen As Long) As String
    If IsInitialized Then Else Initialize
    
    bMap1_SA.pvData = pStr
    bMap1_SA.Bounds.cCount = strLen * 2
    
    VbaMemAllocStringLen = bMap1()
End Function
'аналог SysAllocStringByteLen
Function VbaMemAllocStringByteLen(ByVal pStr As LongPtr, ByVal strBytelen As Long) As String
    If IsInitialized Then Else Initialize
    
    bMap2_SA.pvData = pStr
    bMap2_SA.Bounds.cCount = strBytelen
    
    VbaMemAllocStringByteLen = bMap2()
End Function

'>>>>>>>ARRAY FUNCTIONS<<<<<<<<<<
Private Sub Example_ShellSortS()
    Dim sAr$()
    
    sAr = Split("яблоки Груши аппельсины Кориандр манго")
    
    ShellSortS sAr, Descending, vbTextCompare
End Sub
'http://www.excelworld.ru/board/vba/tricks/sort_array_shell/9-1-0-32
Sub ShellSortS(Arr() As String, _
    Optional ByVal Order As SortOrder = Ascending, Optional ByVal Comp As VbCompareMethod)
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


'>>>>>>>>>>>TESTS<<<<<<<<<<<<<
Private Sub Test_VbaMemAllocStringLen()
    Dim s1$, s2$, s3$
    
    s1 = "df12345da"
    s2 = VbaMemAllocStringLen(StrPtr(s1) + 4, 4) '1234
    s3 = VbaMemAllocStringByteLen(StrPtr(s1) + 10, 6) '45d
End Sub
Private Sub Test_VbaMemRealloc()
    Dim s$, p As LongPtr
    Initialize
    
    p = VbaMemRealloc(0, 6)
    
    lpRef_SA.pvData = VarPtr(s)
    lpRef(0) = p + 4
End Sub
Private Sub TestAllocFree()
    Dim ptr As LongPtr
    Initialize
    ptr = VbaMemAlloc(2)
    VbaMemFreeString ptr + 4
End Sub
Private Sub Test0MemSize()
    Dim ptr As LongPtr, lsz As LongPtr, heap As LongPtr, lres&
    Dim b() As Byte
    Initialize
    heap = GetProcessHeap
    
    b = vbNullString
    
    ptr = GetPtr(ArrPtr(b))
    saRef_SA.pvData = ptr
    Debug.Print HeapSize(heap, 0, saRef(0).pvData)
'    CoTaskMemFree saRef(0).pvData
    VbaMemFree saRef(0).pvData
    saRef(0).pvData = 0
'    saRef(0).Bounds.cCount = 0
    
'    ptr = VbaMemAlloc(2)
'
'    lres = IsBadReadPtr(ptr, 2)
'    lsz = HeapSize(heap, 0, ptr)
'    lres = IsBadReadPtr(ptr, 2)
    
'    VbaMemFree2 ptr
End Sub
Private Sub Example_Ref_Making()
    Dim lp As LongPtr, refDesc As SAFEARRAY1D, ref() As LongPtr
    lp = VarPtr(lp)
    ref = RefPtr(refDesc, VarPtr(lp))
'    MakeRef refDesc, VarPtr(refDesc) - ptrSz, ptrSz
'    refDesc.pvData = lp
End Sub
Private Sub Example_MemLSet()
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
Private Sub Example_MovePtr()
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
Private Sub Example_InStrRev2()
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
'    pSA = RefPtr(pSA)
    CopyPtr pSA, ByVal ArrPtr(sAr)
    
    CopyMemory SA, ByVal pSA, LenB(SA)
End Sub
Private Sub Test_B3()
    Dim b3Ar(2) As B3, b3Ar2(2) As B3
    Debug.Print LenB(b3Ar(0))
    Debug.Print VarPtr(b3Ar(2)) - VarPtr(b3Ar(1))
End Sub
Private Sub TestArrLink()
    Dim ref&(), s$, SA As SAFEARRAY1D
    Initialize
    
    s = "sdfdasd"
    
    With SA
      .cDims = 1
      .fFeatures = FADF_FIXEDSIZE_AUTO
      .cLocks = 1
      .cbElements = 4
      .Bounds.cCount = 1
      .pvData = StrPtr(s) - 4
    End With
    
    lpRef_SA.pvData = VarPtr(s) + ptrSz
    lpRef(0) = VarPtr(SA)
End Sub
Private Sub Test_VariantUnion()
    Dim vs, vbl, vlp, vd
    Initialize
    
    vs = "sdafasfdas"
    vbl = True
    vlp = VarPtr(vlp)
    vd = 3343.0809
    
'    vRef_SA.pvData = VarPtr(vs)
'    Debug.Print vsRef(0).val
    vRef_SA.pvData = VarPtr(vbl)
    Debug.Print viRef(0).val
    vRef_SA.pvData = VarPtr(vlp)
    Debug.Print vlpRef(0).val
    vRef_SA.pvData = VarPtr(vd)
    Debug.Print vdRef(0).val
End Sub
Private Sub Test_VariantUnion2()
    Dim pvArr As LongPtr, vArr(), _
        vsArr() As sVariant, vlpArr() As lpVariant, vdArr() As dVariant, viArr() As iVariant
    Dim i&, SA As SAFEARRAY1D, sRef$(), sRefSA As SAFEARRAY1D
    Initialize
    
    vArr = Array("строка", 344.887, "строка2", True, 323252)
    
    pvArr = VarPtr(pvArr) - ptrSz
    SA = GetSA(GetPtr(pvArr))
    SA.fFeatures = FADF_FIXEDSIZE_AUTO 'FADF_STATIC Or FADF_FIXEDSIZE
    SA.cLocks = 1
'    PutPtr(pvArr - ptrSz) = VarPtr(SA)     'vsArr
    PutPtr(pvArr - ptrSz * 2) = VarPtr(SA)
    PutPtr(pvArr - ptrSz * 3) = VarPtr(SA)
    PutPtr(pvArr - ptrSz * 4) = VarPtr(SA)
    
    sRef = RefStr(sRefSA)
    For i = 0 To UBound(vArr)
        Select Case vlpArr(i).vt
        Case vbDouble, vbDate
            Debug.Print vdArr(i).val
        Case vbString
            sRefSA.pvData = VarPtr(vlpArr(i).val)
            Debug.Print sRef(0)
        Case vbLong
            Debug.Print vlpArr(i).val
        Case vbBoolean, vbInteger
            Debug.Print viArr(i).val
        End Select
    Next
    
'    PutPtr(pvArr - ptrSz) = 0 'требуется освободить vsArr()
End Sub
Private Sub Test_RefStr()
    Dim s1$, ref$(), SA As SAFEARRAY1D, emp$()
    Initialize
    
    s1 = "asfddafa"
    ref = RefStr(SA, VarPtr(s1))
'    Erase ref
    Debug.Print ref(0)
End Sub
