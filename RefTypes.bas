''Attribute VB_Name = "RefTypes"
''================================================================================================================================'
'' RefTypes                                                                                                                       '
''--------------------------------------------------------                                                                        '
'' https://github.com/WNKLER/RefTypes                                                                                             '
''--------------------------------------------------------                                                                        '
'' A VBA/VB6 Library for reading/writing intrinsic types at arbitrary memory addresses.                                           '
'' Its defining feature is that this is achieved using truly native, built-in language features.                                  '
'' It uses no API declarations and has no external dependencies.                                                                  '
''================================================================================================================================'
'' MIT License                                                                                                                    '
''                                                                                                                                '
'' Copyright (c) 2025 Benjamin Dovidio (WNKLER)                                                                                   '
'' Edited by Alexey Leonov (testuser2) 08.2025
'Option Private Module
Option Explicit

#Const SafeMode = True
'#Const PreInitMode = True
Private Enum Context
#If Win64 Then
    [_]   '// 0 on x64; undefined on x86.
#End If
    Win64 '// 0 when [_] is undefined; otherwise, [_] + 1.
    ptrSz = 4 + (4 * Win64)
    varSz = 8 + (2 * ptrSz)
End Enum
Private Const FADF_AUTO      As Integer = &H1
Private Const FADF_FIXEDSIZE As Integer = &H10
#If VBA7 = 0 Then
    Private Enum LONG_PTR: [_]: End Enum
     Public Enum LongPtr:  [_]: End Enum '// Must be Public for Enum-typed Public Property
#End If
#If Not Win64 Then
    Public Type LongLong
        L0x0 As Long
        L0x4 As Long
    End Type
#End If
Private Type HalfPtr
    HalfPtr(-Win64 To 0) As Integer
End Type
Private Type Initializer
    Initializer(-1 To 0) As HalfPtr
End Type
'Private Type Vector          '// Declare Static Fields individually. This avoids defining
'    Element()  As Any        '// a separate Vector UDT for each `Element()` Type.
'    Descriptor As SA1D
'End Type

Public Type SA1D '(SAFEARRAY1D)
    Dims          As Integer
    Features      As Integer
    cbElem        As Long
    Locks         As Long
  #If Win64 Then
    padding       As Long
  #End If
    pData         As LongPtr
    Count         As Long
    lBound        As Long
End Type
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

'##################################################################'
' <Memory proxied by `Initializer`>
Public lpRef() As LongPtr, lpRef_SA As SA1D
Public lpRef2() As LongPtr, lpRef2_SA As SA1D
Public iRef() As Integer, iRef_SA As SA1D
Public iRef2() As Integer, iRef2_SA As SA1D
Public lRef() As Long, lRef_SA As SA1D
Public lRef2() As Long, lRef2_SA As SA1D
Public snRef() As Single, snRef_SA As SA1D
Public dRef() As Double, dRef_SA As SA1D
Public cRef() As Currency, cRef_SA As SA1D
Public cRef2() As Currency, cRef2_SA As SA1D
Public dtRef() As Date, dtRef_SA As SA1D
Public sRef() As String, sRef_SA As SA1D
Public sRef2() As String, sRef2_SA As SA1D
Public oRef() As Object, oRef_SA As SA1D
Public blRef() As Boolean, blRef_SA As SA1D
Public vRef() As Variant, vRef_SA As SA1D
'Union tVariant{
Public vlpRef() As lpVariant
Public vdRef() As dVariant
Public vdtRef() As dtVariant
Public viRef() As iVariant
'Public vsRef() As sVariant '.Locks = 1 does not work
'End Union
Public vRef2() As Variant, vRef2_SA As SA1D
Public unkRef() As IUnknown, unkRef_SA As SA1D
Public bRef() As Byte, bRef_SA As SA1D
Public bRef2() As Byte, bRef2_SA As SA1D
Public llRef() As LongLong, llRef_SA As SA1D
Public iMap1() As Integer, iMap1_SA As SA1D      'мапперы строк (с индексацией от 1)
Public iMap2() As Integer, iMap2_SA As SA1D
Public bMap1() As Byte, bMap1_SA As SA1D
Public bMap2() As Byte, bMap2_SA As SA1D
Public b3Ref1() As B3, b3Ref1_SA As SA1D
Public b3Ref2() As B3, b3Ref2_SA As SA1D         '26
Public saRef() As SA1D, saRef_SA As SA1D  '27
' <End of proxied memory block>
'##################################################################'

Private iMapDyn_SA As SA1D, bMapDyn_SA As SA1D
Public IsInitialized As Boolean, islpRefInit As Boolean

Private Sub MakeRef(Descriptor As SA1D, ByVal ptRef As LongPtr, ByVal cbElem As Long)
    Static This            As Initializer '// Proxy for `Init_Element()`
    Static Init_Element()  As LongPtr     '// Static Init As Vector
    Static Init_Descriptor As SA1D
    
    With Init_Descriptor
        If .Locks Then
        Else
            InitInitializer This.Initializer '// Point `Init_Element()` to `Init_Descriptor`
            .Count = 1
            .Features = FADF_AUTO Or FADF_FIXEDSIZE
            .Dims = 1
            .Locks = 1
            .cbElem = ptrSz
        End If
        
        .pData = ptRef
        Init_Element(0) = VarPtr(Descriptor)
    End With
    
    Descriptor = Init_Descriptor
    Descriptor.cbElem = cbElem
    Descriptor.pData = 0
End Sub
Private Sub InitInitializer(ByRef Initializer() As LONG_PTR)
    Const First As Long = -1
    Const Last  As Long = -0
    
    Initializer(Last) = VarPtr(Initializer(First)) + (2 * ptrSz)
End Sub
Sub Initialize()
    Dim pArr As LongPtr
    If IsInitialized Then Exit Sub
    
    MakeRef lpRef_SA, VarPtr(lpRef_SA) - ptrSz, ptrSz
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
    MakeRef iMap1_SA, VarPtr(iMap1_SA) - ptrSz, 2: iMap1_SA.lBound = 1 'мапперы строк
    MakeRef iMap2_SA, VarPtr(iMap2_SA) - ptrSz, 2: iMap2_SA.lBound = 1
    MakeRef bMap1_SA, VarPtr(bMap1_SA) - ptrSz, 1: bMap1_SA.lBound = 1
    MakeRef bMap2_SA, VarPtr(bMap2_SA) - ptrSz, 1: bMap2_SA.lBound = 1
    MakeRef b3Ref1_SA, VarPtr(b3Ref1_SA) - ptrSz, 3                           'ссылка 3-байтного типа
    MakeRef b3Ref2_SA, VarPtr(b3Ref2_SA) - ptrSz, 3
    MakeRef saRef_SA, VarPtr(saRef_SA) - ptrSz, LenB(saRef_SA) 'ссылка на структуру SafeArray
    
    iMapDyn_SA = iRef_SA: iMapDyn_SA.Locks = 0: iMapDyn_SA.Features = 128
    bMapDyn_SA = bRef_SA: bMapDyn_SA.Locks = 0: bMapDyn_SA.Features = 128
    
    IsInitialized = True
End Sub

Property Get GetInt(ByVal Target As LongPtr) As Integer
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    iRef_SA.pData = Target
    GetInt = iRef(0)
End Property
Property Let PutInt(ByVal Target As LongPtr, ByVal PutInt As Integer)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    iRef_SA.pData = Target
    iRef(0) = PutInt
End Property

Property Get GetLng(ByVal Target As LongPtr) As Long
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lRef_SA.pData = Target
    GetLng = lRef(0)
End Property
Property Let PutLng(ByVal Target As LongPtr, ByVal PutLng As Long)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lRef_SA.pData = Target
    lRef(0) = PutLng
End Property

Property Get GetSng(ByVal Target As LongPtr) As Single
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    snRef_SA.pData = Target
    GetSng = snRef(0)
End Property
Property Let PutSng(ByVal Target As LongPtr, ByVal PutSng As Single)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    snRef_SA.pData = Target
    snRef(0) = PutSng
End Property

Property Get GetDbl(ByVal Target As LongPtr) As Double
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    dRef_SA.pData = Target
    GetDbl = dRef(0)
End Property
Property Let PutDbl(ByVal Target As LongPtr, ByVal PutDbl As Double)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    dRef_SA.pData = Target
    dRef(0) = PutDbl
End Property

Property Get GetCur(ByVal Target As LongPtr) As Currency
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    cRef_SA.pData = Target
    GetCur = cRef(0)
End Property
Property Let PutCur(ByVal Target As LongPtr, ByVal PutCur As Currency)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    cRef_SA.pData = Target
    cRef(0) = PutCur
End Property

Property Get GetDate(ByVal Target As LongPtr) As Date
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    dtRef_SA.pData = Target
    GetDate = dtRef(0)
End Property
Property Let PutDate(ByVal Target As LongPtr, ByVal PutDate As Date)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    dtRef_SA.pData = Target
    dtRef(0) = PutDate
End Property

Property Get GetStr(ByVal Target As LongPtr) As String
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    sRef_SA.pData = Target
    GetStr = sRef(0)
End Property
Property Let PutStr(ByVal Target As LongPtr, ByRef PutStr As String)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    sRef_SA.pData = Target
    sRef(0) = PutStr
End Property
Function RefStr(SA As SA1D, Optional ByVal pData As LongPtr) As String()
    Dim lpArTmp() As String, pTmp As LongPtr
    If islpRefInit Then Else Initialize
    
'    pTmp = lpRef_SA.pData
    SA = sRef_SA
    If pData > 0 Then SA.pData = pData
    lpRef_SA.pData = VarPtr(pTmp) + ptrSz
    lpRef(0) = VarPtr(SA)
'    lpRef_SA.pData = pTmp
    
    RefStr = lpArTmp
End Function

Property Get GetObj(ByVal Target As LongPtr) As Object
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    oRef_SA.pData = Target
    Set GetObj = oRef(0)
End Property
Property Set SetObj(ByVal Target As LongPtr, ByVal SetObj As Object)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    oRef_SA.pData = Target
    Set oRef(0) = SetObj
End Property

Property Get GetBool(ByVal Target As LongPtr) As Boolean
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    blRef_SA.pData = Target
    GetBool = blRef(0)
End Property
Property Let PutBool(ByVal Target As LongPtr, ByVal PutBool As Boolean)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    blRef_SA.pData = Target
    blRef(0) = PutBool
End Property

Property Get GetVar(ByVal Target As LongPtr) As Variant
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    vRef_SA.pData = Target
    GetVar = vRef(0)
End Property
Property Let PutVar(ByVal Target As LongPtr, ByRef PutVar As Variant)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    vRef_SA.pData = Target
    vRef(0) = PutVar
End Property
Property Set SetVar(ByVal Target As LongPtr, ByRef SetVar As Variant)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    vRef_SA.pData = Target
    Set vRef(0) = SetVar
End Property

Property Get GetUnk(ByVal Target As LongPtr) As IUnknown
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    unkRef_SA.pData = Target
    Set RefUnk = unkRef(0)
End Property
Property Set SetUnk(ByVal Target As LongPtr, ByVal RefUnk As IUnknown)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    unkRef_SA.pData = Target
    Set unkRef(0) = RefUnk
End Property

'Property Get GetDec(ByVal Target As LongPtr) As Variant
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
'    dcRef_SA.pData = Target
'    RefDec = dcRef(0)
'End Property
'Property Let PutDec(ByVal Target As LongPtr, ByVal RefDec As Variant)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
'    dcRef_SA.pData = Target
'    dcRef(0) = RefDec
'End Property '_
Property Get GetByte(ByVal Target As LongPtr) As Byte
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    bRef_SA.pData = Target
    RefByte = bRef(0)
End Property
Property Let PutByte(ByVal Target As LongPtr, ByVal RefByte As Byte)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    bRef_SA.pData = Target
    bRef(0) = RefByte
End Property

Property Get GetLngLng(ByVal Target As LongPtr) As LongLong
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
  llRef_SA.pData = Target
  RefLngLng = llRef(0)
End Property
  #If Win64 = 0 Then
Property Let PutLngLng(ByVal Target As LongPtr, ByRef RefLngLng As LongLong)
  #Else
Property Let PutLngLng(ByVal Target As LongPtr, ByVal RefLngLng As LongLong)
  #End If
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
  llRef_SA.pData = Target
  llRef(0) = RefLngLng
End Property

Property Get GetPtr(ByVal Target As LongPtr) As LongPtr
    If islpRefInit Then Else Initialize
    lpRef_SA.pData = Target
    GetPtr = lpRef(0)
End Property
Property Let PutPtr(ByVal Target As LongPtr, ByVal PutPtr As LongPtr)
    If islpRefInit Then Else Initialize
    lpRef_SA.pData = Target
    lpRef(0) = PutPtr
End Property
Function RefPtr(SA As SA1D, Optional ByVal Target As LongPtr) As LongPtr()
    Dim lpArTmp() As LongPtr, pTmp As LongPtr
    If islpRefInit Then Else Initialize
    
'    pTmp = lpRef_SA.pData
    
    SA = lpRef_SA
    SA.pData = Target
    lpRef_SA.pData = VarPtr(pTmp) + ptrSz
    lpRef(0) = VarPtr(SA)
    
'    lpRef_SA.pData = pTmp
    
    RefPtr = lpArTmp
End Function

Property Get GetSA(ByVal Target As LongPtr) As SA1D
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    saRef_SA.pData = Target
    GetSA = saRef(0)
End Property
Property Let PutSA(ByVal Target As LongPtr, RefSA As SA1D)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    saRef_SA.pData = Target
    saRef(0) = PutSA
End Property

Private Sub Test_ArrPtr_()
    Dim bAr() As Byte, iAr%()
    ReDim bAr(2), iAr(2)
    Debug.Print ArrPtr(bAr), GetPtr(ArrPtr(iAr))
    Debug.Print ArrPtrB(bAr), ArrPtrI(iAr, True)
End Sub
Function ArrPtrB(bAry() As Byte, Optional ByVal GetDesc As Boolean) As LongPtr
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lpRef2_SA.pData = VarPtr(GetDesc) - ptrSz
    If GetDesc Then lpRef2_SA.pData = lpRef2(0)
    ArrPtrB = lpRef2(0)
End Function
Function ArrPtrI(iAry() As Integer, Optional ByVal GetDesc As Boolean) As LongPtr
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lpRef2_SA.pData = VarPtr(GetDesc) - ptrSz
    If GetDesc Then lpRef2_SA.pData = lpRef2(0)
    ArrPtrI = lpRef2(0)
End Function

'перемещение указателя (передача владения)
Sub MovePtr(ByVal pDst As LongPtr, ByVal pSrc As LongPtr)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lpRef_SA.pData = pDst
    lpRef2_SA.pData = pSrc
    lpRef(0) = lpRef2(0)
    lpRef2(0) = 0
End Sub
'перемещение указателя строки из Variant в String
Function VarMoveStr(vStr) As String
    If varType(vStr) = vbString Then
      #If Not PreInitMode Then
        If IsInitialized Then Else Initialize
      #End If
        lpRef_SA.pData = VarPtr(VarMoveStr)
        lpRef2_SA.pData = VarPtr(vStr) + 8
        lpRef(0) = lpRef2(0)
        lpRef2(0) = 0
    End If
End Function

Private Sub TestMoveStr()
    Dim s1$, s2$
    s1 = "sdfa"
    s2 = "1122"
    MoveStr s1, s2
End Sub
'безопасное перемещение указателя строки
Sub MoveStr(sDst$, sSrc$)
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    If StrPtr(sDst) Then sDst = vbNullString
    lpRef_SA.pData = VarPtr(sDst)
    lpRef2_SA.pData = VarPtr(sSrc)
    lpRef(0) = lpRef2(0)
    lpRef2(0) = 0
End Sub
'обмен указателями
Sub SwapPtr(ByVal p1 As LongPtr, ByVal p2 As LongPtr)
    Dim pTmp As LongPtr
    lpRef_SA.pData = p1
    lpRef2_SA.pData = p2
    pTmp = lpRef(0)
    lpRef(0) = lpRef2(0)
    lpRef2(0) = pTmp
End Sub

'Аналог CopyMemory
Sub MemLSet(ByVal pDst As LongPtr, ByVal pSrc As LongPtr, ByVal size As Long)
    Dim sDst$, sSrc$, lTmp&
    Dim s1$, s2$
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If size > 3 Then
    Else
        MiniCopy pDst, pSrc, size
        Exit Sub
    End If
    size = size - 4
    
    lRef_SA.pData = pSrc
    lTmp = lRef(0)
    lRef(0) = size
    lRef2_SA.pData = pDst
    lRef2(0) = size

    pSrc = pSrc + 4
    pDst = pDst + 4
    sRef_SA.pData = VarPtr(pSrc)
    sRef2_SA.pData = VarPtr(pDst)
    
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
        bRef_SA.pData = pSrc
        bRef2_SA.pData = pDst
        bRef2(0) = bRef(0)
    ElseIf False Then
2:
        iRef_SA.pData = pSrc
        iRef2_SA.pData = pDst
        iRef2(0) = iRef(0)
    ElseIf False Then
3:
        b3Ref1_SA.pData = pSrc
        b3Ref2_SA.pData = pDst
        b3Ref2(0) = b3Ref1(0)
    End If
End Sub

Function VbaMemRealloc(ByVal pMem As LongPtr, ByVal newSize As Long) As LongPtr
    Dim bMap() As Byte, lp As LongPtr
    If newSize < 1 Then Exit Function
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If pMem Then
    Else
        ReDim bMap(newSize - 1)
        lpRef_SA.pData = VarPtr(lp) + ptrSz
        saRef_SA.pData = lpRef(0)
        VbaMemRealloc = saRef(0).pData ' = VarPtr(bMap(0))
        saRef(0).pData = 0
        Exit Function
    End If
    
    bMapDyn_SA.pData = pMem
    bMapDyn_SA.Count = newSize
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(bMapDyn_SA)
    ReDim Preserve bMap(newSize - 1)
    lpRef(0) = 0
    VbaMemRealloc = bMapDyn_SA.pData
End Function
Function VbaMemAlloc(ByVal size As LongPtr) As LongPtr
    Dim bMap() As Byte, lp As LongPtr
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    ReDim bMap(size - 1)
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    saRef_SA.pData = lpRef(0)
    With saRef(0)
      VbaMemAlloc = .pData ' = VarPtr(bMap(0))
      .pData = 0
    End With
End Function
'Sub VbaMemFree2(ByVal ptr As LongPtr)
'    Dim bMap() As Byte, lp As LongPtr
'    bMap = vbNullString 'Array()
'    lpRef_SA.pData = VarPtr(lp) + ptrSz
'    saRef_SA.pData = lpRef(0)
'    saRef(0).pData = ptr
'End Sub
Sub VbaMemFree(ByVal ptr As LongPtr)
    Dim s$
    lpRef_SA.pData = VarPtr(s)
    lpRef(0) = ptr + 4
End Sub

Function GetBytMap(SA As SA1D, ByVal ptr As LongPtr, ByVal cbCnt As LongPtr) As Byte()
    Dim bMap() As Byte, lp As LongPtr
    SA = bMap2_SA
    SA.pData = ptr
    SA.Count = cbCnt
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(SA)
    GetBytMap = bMap
End Function
Function GetIntMap(SA As SA1D, ByVal ptr As LongPtr, ByVal ciCnt As LongPtr) As Integer()
    Dim iMap() As Integer, lp As LongPtr
    SA = iMap2_SA
    SA.pData = ptr
    SA.Count = ciCnt
    lpRef_SA.pData = VarPtr(lp) + ptrSz
    lpRef(0) = VarPtr(SA)
    GetIntMap = iMap
End Function

'>>>>>>>>>>>>>>>STRINGS SECTION<<<<<<<<<<<<<<<<<<<'
Private Function StrCompVBA(str1$, str2$) As Long
    Dim len1&, len2&, lenMin&
    Dim i&, dif&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    len1 = Len(str1) + 1: len2 = Len(str2) + 1
    If len1 > len2 Then lenMin = len2 Else lenMin = len1
    iRef1_SA.pData = StrPtr(str1)
    iRef1_SA.Count = lenMin
    iRef2_SA.pData = StrPtr(str2)
    iRef2_SA.Count = lenMin
    
    For i = 1 To lenMin
        dif = iRef1(i) - iRef2(i)
        If dif Then Exit For
    Next
    
    StrCompVBA = dif
End Function
'аналог instr$() с дополнителным параметром endFind, чтобы указывать позицию окончания поиска.
Function InStrEnd(sCheck$, sMatch$, Optional ByVal Start As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal endFind As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pData = StrPtr(sCheck)
        iMap2_SA.pData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        sTmp1 = UCase$(sCheck)
        sTmp2 = UCase$(sMatch)
        iMap1_SA.pData = StrPtr(sTmp1)
        iMap2_SA.pData = StrPtr(sTmp2)
    End If
    lenCheck = Len(sCheck)
    lenMatch = Len(sMatch)
    iMap1_SA.Count = lenCheck
    iMap2_SA.Count = lenMatch
    If endFind = -1 Then endFind = lenCheck
    
    iMatch = iMap2(1)                                                   'v2
    For i = Start To endFind - lenMatch + 1
        If iMap1(i) <> iMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If iMap1(k) = iMap2(j) Then Else GoTo skip
            Next
            InStrEnd = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrEndB(sCheck$, sMatch$, Optional ByVal Start As Long = 1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal endFind As Long = -1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If Compare = vbBinaryCompare Then
        bMap1_SA.pData = StrPtr(sCheck)
        bMap2_SA.pData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        sTmp1 = UCase$(sCheck)
        sTmp2 = UCase$(sMatch)
        bMap1_SA.pData = StrPtr(sTmp1)
        bMap2_SA.pData = StrPtr(sTmp2)
    End If
    lenCheck = LenB(sCheck)
    lenMatch = LenB(sMatch)
    bMap1_SA.Count = lenCheck
    bMap2_SA.Count = lenMatch
    If endFind = -1 Then endFind = lenCheck
    
    bMatch = bMap2(1)                                                   'v2
    For i = Start To endFind - lenMatch + 1
        If bMap1(i) <> bMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If bMap1(k) = bMap2(j) Then Else GoTo skip
            Next
            InStrEndB = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrEndRev(sCheck$, sMatch$, Optional ByVal Start As Long = -1, _
    Optional ByVal Compare As VbCompareMethod, Optional ByVal endFind As Long = 1) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, iMatch%
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If Compare = vbBinaryCompare Then
        iMap1_SA.pData = StrPtr(sCheck)
        iMap2_SA.pData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        sTmp1 = UCase$(sCheck)
        sTmp2 = UCase$(sMatch)
        iMap1_SA.pData = StrPtr(sTmp1)
        iMap2_SA.pData = StrPtr(sTmp2)
    End If
    lenCheck = Len(sCheck)
    lenMatch = Len(sMatch)
    iMap1_SA.Count = lenCheck
    iMap2_SA.Count = lenMatch
    If Start = -1 Then Start = lenCheck
    
'    Dim bgnIter& '                                                     'v1
'    j = lenMatch
'    iMatch = iMap2(lenMatch)
'    bgnIter = lenMatch - 1
'    For i = Start To endFind + lenMatch - 1 Step -1
'        If iMap1(i) <> iMatch Then
'        Else
'            k = i
'            For j = bgnIter To 1 Step -1
'                k = k - 1
'                If iMap1(k) = iMap2(j) Then Else GoTo skip
'            Next
'            InStrEndRev = k: Exit Function
'        End If
'skip:
'    Next
    iMatch = iMap2(1)                                                   'v2
    For i = Start - lenMatch + 1 To endFind Step -1
        If iMap1(i) <> iMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If iMap1(k) = iMap2(j) Then Else GoTo skip
            Next
            InStrEndRev = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrEndRevB(sCheck$, sMatch$, Optional ByVal Start As Long = -1, _
    Optional ByVal endFind As Long = 1, Optional ByVal Compare As VbCompareMethod) As Long
    Dim i&, j&, k&, lenCheck&, lenMatch&, bMatch As Byte
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    If Compare = vbBinaryCompare Then
        bMap1_SA.pData = StrPtr(sCheck)
        bMap2_SA.pData = StrPtr(sMatch)
    Else
        Dim sTmp1$, sTmp2$
        sTmp1 = UCase$(sCheck)
        sTmp2 = UCase$(sMatch)
        bMap1_SA.pData = StrPtr(sTmp1)
        bMap2_SA.pData = StrPtr(sTmp2)
    End If
    lenCheck = LenB(sCheck)
    lenMatch = LenB(sMatch)
    bMap1_SA.Count = lenCheck
    bMap2_SA.Count = lenMatch
    If Start = -1 Then Start = lenCheck
    
    bMatch = bMap2(1)                                                   'v2
    For i = Start - lenMatch + 1 To endFind Step -1
        If bMap1(i) <> bMatch Then
        Else
            k = i
            For j = 2 To lenMatch
                k = k + 1
                If bMap1(k) = bMap2(j) Then Else GoTo skip
            Next
            InStrEndRevB = i: Exit Function
        End If
skip:
    Next
End Function
Function InStrLenRevB(sCheck$, sMatch$, Optional ByVal Start As Long = -1, _
    Optional ByVal Length As Long = 1, Optional ByVal Compare As VbCompareMethod) As Long
    If Start = -1 Then Start = LenB(sCheck)
    InStrLenRevB = InStrEndRevB(sCheck, sMatch, Start, Start - Length + 1, Compare)
End Function

Function InStrLen(ByVal Start As Long, sCheck$, sMatch$, ByVal lenFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    Dim szCheck&, newSize&, pCheck As LongPtr
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    pCheck = StrPtr(sCheck)
    If pCheck Then Else Exit Function
    
    lRef_SA.pData = pCheck - 4
    szCheck = lRef(0)
    newSize = (Start + lenFind - 1) * 2
    If newSize < szCheck Then
        If newSize > -1 Then Else GoTo errArgum
        lRef(0) = newSize
        InStrLen = InStr(Start, sCheck, sMatch, Compare)
        lRef(0) = szCheck
    Else: InStrLen = InStr(Start, sCheck, sMatch, Compare)
    End If
    
    Exit Function
errArgum:
    Err.Raise 5, , "invalid function argumenct"
End Function
Function InStrLenB(ByVal Start As Long, sCheck$, sMatch$, ByVal lenFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    Dim szCheck&, newSize&, pCheck As LongPtr
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    pCheck = StrPtr(sCheck)
    If pCheck Then Else Exit Function
    
    lRef_SA.pData = pCheck - 4
    szCheck = lRef(0)
    newSize = Start + lenFind - 1
    If newSize < szCheck Then
        If newSize > -1 Then Else GoTo errArgum
        lRef(0) = newSize
        InStrLenB = InStrB(Start, sCheck, sMatch, Compare)
        lRef(0) = szCheck
    Else: InStrLenB = InStrB(Start, sCheck, sMatch, Compare)
    End If
    
    Exit Function
errArgum:
    Err.Raise 5, , "invalid function argumenct"
End Function
'No safe version without any checks NS = Not Safe

Function InStrEndNS(ByVal Start As Long, sCheck$, sMatch$, ByVal endFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    Dim lTmp&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lRef_SA.pData = StrPtr(sCheck) - 4
    lTmp = lRef(0)
    
    lRef(0) = endFind * 2
    InStrEndNS = InStr(Start, sCheck, sMatch, Compare)
    lRef(0) = lTmp
End Function
Function InStrLenNS(ByVal Start As Long, sCheck$, sMatch$, ByVal lenFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    InStrLenNS = InStrEndNS(Start, sCheck, sMatch, Start + lenFind - 1, Compare)
End Function

Function InStrEndBNS(ByVal Start As Long, sCheck$, sMatch$, ByVal endFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    Dim lTmp&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    lRef_SA.pData = StrPtr(sCheck) - 4
    lTmp = lRef(0)
    
    lRef(0) = endFind
    InStrEndBNS = InStrB(Start, sCheck, sMatch, Compare)
    lRef(0) = lTmp
End Function
Function InStrLenBNS(ByVal Start As Long, sCheck$, sMatch$, ByVal lenFind As Long, _
    Optional ByVal Compare As VbCompareMethod) As Long
    InStrLenBNS = InStrEndBNS(Start, sCheck, sMatch, Start + lenFind - 1, Compare)
End Function

Function InStrEndRevNS(sCheck$, sMatch$, ByVal Start As Long, _
    Optional ByVal endFind As Long = 1, Optional ByVal Compare As VbCompareMethod) As Long
    Dim sTmp$, pTmp As LongPtr, lTmp&, lRet&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    endFind = endFind - 1
    pTmp = StrPtr(sCheck) + endFind * 2
    lRef_SA.pData = pTmp - 4
    lTmp = lRef(0)
    lRef(0) = (Start - endFind) * 2
    lpRef_SA.pData = VarPtr(sTmp)
    lpRef(0) = pTmp
    
    lRet = InStrRev(sTmp, sMatch, , Compare)
    If lRet Then InStrEndRevNS = lRet + endFind
    
    lpRef(0) = 0
    lRef(0) = lTmp
End Function
Function InStrLenRevNS(sCheck$, sMatch$, ByVal Start As Long, _
    Optional ByVal lenFind As Long = -1, Optional ByVal Compare As VbCompareMethod) As Long
    Dim endFind&: endFind = Start - lenFind + 1
    InStrLenRevNS = InStrEndRevNS(sCheck, sMatch, Start, endFind, Compare)
End Function
Private Sub Test_InStrLen()
    Dim s1$, s2$
    Dim l1&, l2&, l3&, l4&
    Initialize
    s1 = "dretilk';nnll8"
    s2 = "tilk"
    
'    l1 = InStrLen(8, s1, s2, 4)
'    l2 = InStrLenB(15, s1, s2, 8)
'    lRef_SA.pData = StrPtr(s1) - 4
'    l3 = InStrLenNS(8, s1, s2, 4, , lRef(0))
'    l4 = InStrLenBNS(15, s1, s2, 8, , lRef(0))
    l1 = InStrRev(s1, s2, 11)
    l2 = InStrEndRevNS(s1, s2, 7, 4)
    l3 = InStrLenRevNS(s1, s2, 7, 4)
End Sub
Sub TestProxyRef(Optional l&, Optional ByVal l0&)
    Dim s$
    Initialize
    s = "sdfdfsadf"
'    lRef_SA.pData = StrPtr(s) - 4
    PutPtr(VarPtr(l0) - ptrSz) = StrPtr(s) - 4
    l = 8
    TestProxyRef_ s, l 'lRef(0)
End Sub
Private Sub TestProxyRef_(s$, l&)
    l = 9
End Sub
Private Sub TestResizeString()
    Dim s$: s = "abcd"
    
    ReallocStringB s, 11
End Sub
Sub ReallocString(sSrc$, ByVal newSize&)
    Dim iMap%(), pSrc As LongPtr
    If newSize < 0 Then Exit Sub
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
        
    pSrc = StrPtr(sSrc)
    If pSrc Then
        iMapDyn_SA.pData = StrPtr(sSrc) - 4
        iMapDyn_SA.Count = Len(sSrc) + 3
    Else
        sSrc = String$(newSize, vbNullChar)
        Exit Sub
    End If
        
    lpRef_SA.pData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(iMapDyn_SA)
    ReDim Preserve iMap(newSize + 2)
    lpRef(0) = 0
    
    lpRef2_SA.pData = VarPtr(sSrc)
    lpRef2(0) = iMapDyn_SA.pData + 4
    lRef_SA.pData = iMapDyn_SA.pData
    lRef(0) = newSize * 2
End Sub
Sub ReallocStringB(sSrc$, ByVal newSize&)
    Dim bMap() As Byte, pSrc As LongPtr
    If newSize < 0 Then Exit Sub
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
        
    pSrc = StrPtr(sSrc)
    If pSrc Then
    Else
        ReDim bMap(newSize - 1) As Byte
        sSrc = bMap
        Exit Sub
    End If
    bMapDyn_SA.pData = StrPtr(sSrc) - 4
    bMapDyn_SA.Count = LenB(sSrc) + 6
        
    lpRef_SA.pData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(bMapDyn_SA)
    
    ReDim Preserve bMap(newSize + 5)
    lpRef(0) = 0
    
    lpRef2_SA.pData = VarPtr(sSrc)
    lpRef2(0) = bMapDyn_SA.pData + 4
    
    lRef_SA.pData = bMapDyn_SA.pData
    lRef(0) = newSize
End Sub
'аналог SysAllocStringLen Функции копирующей заданный буфер в новую bstr-строку
Function VbaMemAllocStringLen(ByVal pStr As LongPtr, ByVal strlen As Long) As String
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    bMap1_SA.pData = pStr
    bMap1_SA.Count = strlen * 2
    
    VbaMemAllocStringLen = bMap1()
End Function
'аналог SysAllocStringByteLen
Function VbaMemAllocStringByteLen(ByVal pStr As LongPtr, ByVal strBytelen As Long) As String
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    bMap2_SA.pData = pStr
    bMap2_SA.Count = strBytelen
    
    VbaMemAllocStringByteLen = bMap2()
End Function
Function StartsWith(sCheck$, sMatch$) As Boolean
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    Dim lTmp&, szMatch&
    lRef_SA.pData = StrPtr(sCheck) - 4
    lTmp = lRef(0)
    szMatch = LenB(sMatch)
    If lTmp >= szMatch Then
        lRef(0) = szMatch
        StartsWith = (sCheck = sMatch)
        lRef(0) = lTmp
    End If
End Function
Function EndsWith(sCheck$, sMatch$) As Boolean
    Dim szCheck&, szMatch&                                    'v1 (no ref. used)
    szCheck = LenB(sCheck)
    szMatch = LenB(sMatch)
    EndsWith = InStrB(szCheck - szMatch + 1, sCheck, sMatch)
'    EndsWith = InStrLenRevB(sCheck, sMatch, , , LenB(sMatch)) 'v2
End Function
'no ref. used
Function Repeat(ByVal Count&, sSrc$) As String
'    Dim lnSrc&, lnRes&, i&                     'v1
'    lnSrc = Len(sSrc)
'    lnRes = lnSrc * Count
'    Repeat = VarMoveStr(String(lnRes, vbNullChar))
'    For i = 1 To lnRes - lnSrc + 1 Step lnSrc
'        Mid$(Repeat, i, (lnSrc)) = sSrc
'    Next
    Dim pDst As LongPtr, pSrc As LongPtr, szSrc& 'v2
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    pSrc = StrPtr(sSrc)
    szSrc = LenB(sSrc)
    Repeat = VarMoveStr(String((szSrc \ 2) * Count, vbNullChar))
    pDst = StrPtr(Repeat)
    For pDst = pDst To pDst + szSrc * (Count - 1) Step szSrc
        MemLSet pDst, pSrc, szSrc
    Next
End Function
Function StringB(ByVal Num As Long, Char) As String
    Dim bChar As Byte, i&, bBuf() As Byte
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    If varType(Char) = vbString Then
        bChar = Asc(Char)
    ElseIf IsNumeric(Char) Then
        If Char > -1 Then Else Exit Function
        bChar = Char
    Else: Exit Function
    End If
    ReDim bBuf(Num + 5)
    If bChar Then
        For i = 4 To 3 + Num
            bBuf(i) = bChar
        Next
    End If
    lpRef_SA.pData = VarPtr(i) - ptrSz
    saRef_SA.pData = lpRef(0)
    lRef_SA.pData = saRef(0).pData ' VarPtr(bBuf(0))
    lRef(0) = Num
    lpRef_SA.pData = VarPtr(StringB)
    lpRef(0) = saRef(0).pData + 4
'    saRef(0).Count = 0
    saRef(0).pData = 0
End Function
Function Strip(sSource$, Optional sTrimChars$) As String
    Static init As Boolean, dfltTrimDict As New Dictionary, curTrimDict As New Dictionary
    Dim i&, lOff&, rOff&, Desc2 As SA1D, iTrimChars%(), Desc1 As SA1D, iSource%()
    Dim lnSrc&, cTrimChars&, TrimDict As Dictionary
    If init Then
    Else
        dfltTrimDict.Add 9, Empty 'vbTab
        dfltTrimDict.Add 10, Empty 'vbLf
        dfltTrimDict.Add 11, Empty '\v
        dfltTrimDict.Add 12, Empty 'vbFormFeed
        dfltTrimDict.Add 13, Empty 'vbCr/vbNewLine
        dfltTrimDict.Add 32, Empty '" "
        init = True
    End If
    lnSrc = Len(sSource)
    If lnSrc Then Else Exit Function
    
    iSource = GetStrMap(Desc1, sSource)
    cTrimChars = Len(sTrimChars)
    If cTrimChars Then
        iTrimChars = GetStrMap(Desc2, sTrimChars)
        For i = 1 To cTrimChars
            curTrimDict(iTrimChars(i)) = Empty
        Next
        Set TrimDict = curTrimDict
    Else
        Set TrimDict = dfltTrimDict
    End If
    
    For i = 1 To lnSrc
        If TrimDict.Exists(iSource(i)) Then
            lOff = lOff + 1
        Else: Exit For
        End If
    Next
    For i = lnSrc To lOff + 1 Step -1
        If TrimDict.Exists(iSource(i)) Then
            rOff = rOff + 1
        Else: Exit For
        End If
    Next
    If cTrimChars Then curTrimDict.RemoveAll
    
    Strip = Mid$(sSource, 1 + lOff, lnSrc - lOff - rOff)
    ''for StripBuf()
'    Dim pSrc As LongPtr, pTmp As LongPtr, szNew&
'    If lOff Then
'        pSrc = StrPtr(sSource)
'        pTmp = pSrc + 2 * lOff
'        szNew = (lnSrc - lOff) * 2
'
'        lRef_SA.pData = pTmp - 4
'        lRef(0) = szNew
'        lRef2_SA.pData = pSrc - 4
'        lRef2(0) = szNew
'        sRef_SA.pData = VarPtr(pTmp)
'        LSet sSource = sRef(0)
''        sRef_SA.pData = 0
'        lRef2(0) = lnSrc * 2
'    End If
'    If lOff > 0 Or rOff > 0 Then
'        ReallocString sSource, lnSrc - lOff - rOff
'    End If
End Function
Private Sub Test_InsertBuf()
    Dim s1$, s2$, s3$
    Initialize
    s1 = "sdkjirpoiw"
    s2 = "1122"
    
'    InsertBuf s1, 11, s2
    s3 = Insert(s1, 12, s2)
    Debug.Print s3 's1       'sdkj1122irpoiw
End Sub
Function Insert(sSrc$, ByVal pos&, sIns$) As String
    Dim lnSrc&, lnIns&, lnRes&, szTmp&, lTmp&
    Dim sRes$, pSrc As LongPtr, pTmp1 As LongPtr, pTmp2 As LongPtr
    
    lnSrc = Len(sSrc): lnIns = Len(sIns)
    If lnSrc > 0 And lnIns > 0 Then Else Exit Function
    If pos < 0 Then
        pos = lnSrc + pos + 1
    End If
    Select Case pos
    Case 1 To lnSrc + 1
    Case Else: Exit Function
    End Select
    
    lnRes = lnSrc + lnIns
    pSrc = StrPtr(sSrc)
    sRes = VarMoveStr(String(lnRes, vbNullChar))
    pTmp1 = StrPtr(sRes)
    szTmp = (pos - 1) * 2
    If szTmp Then
        lRef_SA.pData = pSrc - 4:
        lRef2_SA.pData = pTmp1 - 4
        
        lRef(0) = szTmp: lRef2(0) = szTmp
        LSet sRes = sSrc
        lRef(0) = lnSrc * 2: lRef2(0) = lnRes * 2
    End If
    pTmp2 = pSrc + szTmp
    
    pTmp1 = pTmp1 + szTmp
    szTmp = lnIns * 2
    lRef_SA.pData = pTmp1 - 4
    lTmp = lRef(0)
    lRef(0) = szTmp
    sRef_SA.pData = VarPtr(pTmp1)
    LSet sRef(0) = sIns
'    sRef_SA.pData = 0
    lRef(0) = lTmp
    
    If pos <= lnSrc Then
        pTmp1 = pTmp1 + szTmp
        szTmp = (lnSrc - pos + 1) * 2
        MemLSet pTmp1, pTmp2, szTmp
    End If
    
    MovePtr VarPtr(Insert), VarPtr(sRes)
End Function
Sub InsertBuf(sSrc$, ByVal pos&, sIns$)
    Dim lnSrc&, lnIns&, lnRes&, szTmp&, szIns&, lTmp1&, lTmp2&
    Dim pSrc As LongPtr, pTmp1 As LongPtr, pTmp2 As LongPtr
    
    lnSrc = Len(sSrc): lnIns = Len(sIns)
    If lnSrc > 0 And lnIns > 0 Then Else Exit Sub
    If pos < 0 Then
        pos = lnSrc + pos + 1
    End If
    Select Case pos
    Case 1 To lnSrc + 1
    Case Else: Exit Function
    End Select
    
    lnRes = lnSrc + lnIns
    ReallocString sSrc, lnRes
    pSrc = StrPtr(sSrc)
    pTmp1 = pSrc + (pos - 1) * 2
    szIns = lnIns * 2
    pTmp2 = pTmp1 + szIns
    szTmp = (lnSrc - pos + 1) * 2
    lRef_SA.pData = pTmp1 - 4
    lTmp1 = lRef(0)
    lRef(0) = szTmp
    lRef2_SA.pData = pTmp2 - 4
    lTmp2 = lRef2(0)
    lRef2(0) = szTmp
    sRef_SA.pData = VarPtr(pTmp1)
    sRef2_SA.pData = VarPtr(pTmp2)
    LSet sRef2(0) = sRef(0)
    
    lRef2_SA.pData = lRef2_SA.pData + szIns
    lRef2(0) = lTmp2
    
    lRef(0) = szIns
    LSet sRef(0) = sIns
    lRef(0) = lTmp1
End Sub
Private Sub Test_Remove()
    Dim s$, s2$
    s = "abcdefgh"
    Debug.Print Remove(s, 1, 4)   'efgh
    Debug.Print Remove(s, 5, -2)   'abcfgh
    Debug.Print Remove(s, -3, -3) 'abcgh
    
End Sub
Function Remove(sSrc$, ByVal Start&, ByVal Length&) As String
    Dim lnSrc&, maxLen&
    Dim sRes$, lnRes&, lnTmp&
    Dim pSrc As LongPtr, pDst As LongPtr, szTmp&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
  
    lnSrc = Len(sSrc)
    If lnSrc Then Else Exit Function
    If Start > 0 Then
    ElseIf Start < 0 Then
        Start = lnSrc + Start + 1
    End If
    If Length > 0 Then
    ElseIf Length < 0 Then
        Start = Start + Length + 1
        If Start > 0 Then Else Exit Function
        Length = -Length
    Else: Exit Function
    End If
    If Start <= lnSrc Then Else Exit Function
    maxLen = lnSrc - Start + 1
    If Length > maxLen Then Length = maxLen
    
    lnRes = lnSrc - Length
    sRes = String(lnRes, vbNullChar)
    pSrc = StrPtr(sSrc): pDst = StrPtr(sRes)
    If Start > 1 Then
'        lnTmp = Start - 1
'        Mid$(sRes, 1, (lnTmp)) = Mid$(sSrc, 1, (lnTmp))       'v1
        szTmp = (Start - 1) * 2                               'v2
        lRef_SA.pData = pDst - 4
        lRef(0) = szTmp
        lRef2_SA.pData = pSrc - 4
        lRef2(0) = szTmp
        LSet sRes = sSrc
        lRef(0) = lnRes * 2
        lRef2(0) = lnSrc * 2
    End If
    If Length < maxLen Then
'        lnTmp = maxLen - Length
'        Mid$(sRes, Start, (lnTmp)) = Mid$(sSrc, Start + Length, (lnTmp)) 'v1
        szTmp = (maxLen - Length) * 2                          'v2
        pDst = pDst + (Start - 1) * 2
        pSrc = pSrc + (Start + Length - 1) * 2
        MemLSet pDst, pSrc, szTmp
    End If
    
    MoveStr Remove, sRes
End Function
Function padCenter(sSrc$, ByVal Length&, Optional Char) As String
    Dim iMap%(), pSrc As LongPtr, pDst As LongPtr, lnSrc&, lOff&, lTmp&
    Dim i&, iChar%, pDat As LongPtr
    
    lnSrc = Len(sSrc)
    If Length > lnSrc Then Else Exit Function
    If varType(Char) = vbString Then
        iChar = AscW(Char)
    ElseIf IsNumeric(Char) Then
        iChar = Char
    Else: iChar = 32
    End If
    
    lOff = (Length - lnSrc) / 2 + 1
    padCenter = VarMoveStr(String(Length, iChar))
    Mid$(padCenter, lOff, (lnSrc)) = sSrc
End Function
Function padStart(sSrc$, ByVal Length&, Optional Char) As String
    Dim iMap%(), pSrc As LongPtr, pDst As LongPtr, lnSrc&, lOff&, lTmp&
    Dim i&, iChar%, pDat As LongPtr
    
    lnSrc = Len(sSrc)
    If Length > lnSrc Then Else Exit Function
    If varType(Char) = vbString Then
        iChar = AscW(Char)
    ElseIf IsNumeric(Char) Then
        iChar = Char
    Else: iChar = 32
    End If

    pDat = StrPtr(sSrc) - 4
    iMapDyn_SA.pData = pDat
    iMapDyn_SA.Count = lnSrc + 3
    lpRef_SA.pData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(iMapDyn_SA)
      ReDim Preserve iMap(Length + 2)
      
      pSrc = iMapDyn_SA.pData + 4
      If pDat <> iMapDyn_SA.pData Then
          lpRef2_SA.pData = VarPtr(sSrc)
          lpRef2(0) = pSrc
      End If
      lRef_SA.pData = iMapDyn_SA.pData
      lRef(0) = Length * 2
      
      lOff = Length - lnSrc
      pDst = pSrc + (lOff) * 2
      
      lRef_SA.pData = pDst - 4
      lTmp = lRef(0)
      lRef(0) = lnSrc * 2
      sRef_SA.pData = VarPtr(pDst)
      
      LSet sRef(0) = sSrc
      
      lRef_SA.pData = pDst + (lOff) * 2 - 4
      lRef(0) = lTmp
      
      For i = 2 To lOff + 1
          iMap(i) = iChar
      Next
    lpRef(0) = 0
End Function
Function padEnd(sSrc$, ByVal Length&, Optional Char) As String
    Dim iMap%(), lnSrc&
    Dim i&
    
    lnSrc = Len(sSrc)
    If Length > lnSrc Then Else Exit Function
    
    padEnd = VarMoveStr(String(Length, Char))
    Mid$(padEnd, 1, (lnSrc)) = sSrc
End Function
Sub padEndBuf(sSrc$, ByVal Length&, Optional Char)
    Dim iMap%(), pSrc As LongPtr, lnSrc&
    Dim i&, iChar%, pDat As LongPtr
    
    lnSrc = Len(sSrc)
    If Length > lnSrc Then Else Exit Sub
    If varType(Char) = vbString Then
        iChar = AscW(Char)
    ElseIf IsNumeric(Char) Then
        iChar = Char
    Else: iChar = 32
    End If
    
    pDat = StrPtr(sSrc) - 4
    iMapDyn_SA.pData = pDat
    iMapDyn_SA.Count = lnSrc + 3
    lpRef_SA.pData = VarPtr(pSrc) + ptrSz
    lpRef(0) = VarPtr(iMapDyn_SA)
      ReDim Preserve iMap(Length + 2)
      
      pSrc = iMapDyn_SA.pData + 4
      If pDat <> iMapDyn_SA.pData Then
          lpRef2_SA.pData = VarPtr(sSrc)
          lpRef2(0) = pSrc
      End If
      lRef_SA.pData = iMapDyn_SA.pData
      lRef(0) = Length * 2
      
      For i = lnSrc + 2 To Length + 1
          iMap(i) = iChar
      Next
    lpRef(0) = 0
End Sub

'>>>>>>>ARRAY FUNCTIONS<<<<<<<<<<
'http://www.excelworld.ru/board/vba/tricks/sort_array_shell/9-1-0-32
Sub ShellSortS(Arr() As String, _
    Optional ByVal Order As SortOrder = Ascending, Optional ByVal Comp As VbCompareMethod)
    Dim Limit&, Switch&, i&, j&, ij&, ub&
  #If Not PreInitMode Then
    If IsInitialized Then Else Initialize
  #End If
    
    ub = UBound(Arr)
    j = (ub + 1) \ 2
    Do While j > 0
        Limit = ub - j
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
Private Sub Test_padStart()
    Dim s1$, s2$, s3$
    Initialize
    s1 = "12345"
'    s2 = padStart(s1, 10, 46)
'    s2 = padEnd(s1, 8, "0")
    s3 = padCenter(s1, 12, "|")
    Debug.Print s3
End Sub
Private Sub Test_Strip()
    Dim s$, s2$
    s = "  #*Hello World!..-  "
    s2 = Strip(s, " #*-")
End Sub
Private Sub Test_StringB()
    Dim s$, s2$
    
    s = StringB(10, vbNullChar)
    s2 = String(10, 0)
End Sub
Private Sub Test_InStrLenRevB()
    Dim s1$, s2$, lres1&, lres2&
    s1 = "kdsFddLjklk" '22-10 13
    s2 = "jklk"
'    lres1 = InStrRev(s1, s2)
    lres1 = InStrEndRevB(s1, s2, , 15)
    lres2 = InStrLenRevB(s1, s2, , 8)
End Sub
Private Sub TestVarMoveString()
    Dim v, s$
    v = "sdfdasfdas"
    s = VarMoveStr(v)
End Sub
Private Sub Test_Repeat()
    Dim s$, s2$
    s = "ha"
    s2 = Repeat(3, s)
End Sub
Private Sub Test_StartsWith_EndsWith()
    Dim s1$, s2$, bl1 As Boolean, bl2 As Boolean
    s1 = "телевизор"
    
    bl1 = StartsWith(s1, "тел")
    bl2 = EndsWith(s1, "изор")
End Sub
Private Sub Example_ShellSortS()
    Dim sAr$()
    
    sAr = Split("яблоки Груши аппельсины Кориандр манго")
    
    ShellSortS sAr, Descending, vbTextCompare
End Sub
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
    
    lpRef_SA.pData = VarPtr(s)
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
    saRef_SA.pData = ptr
    Debug.Print HeapSize(heap, 0, saRef(0).pData)
'    CoTaskMemFree saRef(0).pData
    VbaMemFree saRef(0).pData
    saRef(0).pData = 0
'    saRef(0).Count = 0
    
'    ptr = VbaMemAlloc(2)
'
'    lres = IsBadReadPtr(ptr, 2)
'    lsz = HeapSize(heap, 0, ptr)
'    lres = IsBadReadPtr(ptr, 2)
    
'    VbaMemFree2 ptr
End Sub
Private Sub Example_Ref_Making()
    Dim lp As LongPtr, refDesc As SA1D, ref() As LongPtr
    lp = VarPtr(lp)
    ref = RefPtr(refDesc, VarPtr(lp))
'    MakeRef refDesc, VarPtr(refDesc) - ptrSz, ptrSz
'    refDesc.pData = lp
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
Private Sub Example_InStrEndRev()
    Dim sCheck$, sMatch$, lres&, lres2&, cmp As VbCompareMethod
    sCheck = "rtoiutPoIpkj"
    sMatch = "TpoI"
    cmp = TextCompare
    lres = InStrEndRev(sCheck, sMatch, 9, vbTextCompare, 6)
    lres2 = InStrEndRevB(sCheck, sMatch, 18, vbTextCompare, 11)
    lres = InStrEnd(sCheck, sMatch, 6, cmp, 9)
    lres2 = InStrEndB(sCheck, sMatch, 11, cmp, 18)
    Stop
End Sub
Private Sub TestiRef()
    Dim s$: s = "АБВ"
    Initialize
    iRef_SA.pData = StrPtr(s)
    iRef_SA.Count = Len(s)
    iRef(2) = AscW("Ъ")
    ReDim Preserve iRef(1 To 3)
End Sub
Private Sub TestArrayCopy()
    Dim s1$, s2$
    s1 = "АБВГД"
    s2 = "     "
    Initialize
    With iRef_SA
      .pData = StrPtr(s1)
      .Count = Len(s1)
    End With
    With m_SA
      .pData = StrPtr(s2)
      .cbElem = 2
      .lBound = 1
      .Count = Len(s2)
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
    iRef1_SA.pData = StrPtr(s1)
    iRef1_SA.Count = Len(s1) + 1
    iRef2_SA.pData = StrPtr(s2)
    iRef2_SA.Count = Len(s1) + 1
    iRef2(7) = 123
    LSet s2 = s1
    Debug.Print iRef2(7)
End Sub
Private Sub TestArray()
    Dim sAr$(2), pSA As LongPtr, SA As SA1D
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
    Dim ref&(), s$, SA As SA1D
    Initialize
    
    s = "sdfdasd"
    
    With SA
      .Dims = 1
      .Features = FADF_FIXEDSIZE_AUTO
      .Locks = 1
      .cbElem = 4
      .Count = 1
      .pData = StrPtr(s) - 4
    End With
    
    lpRef_SA.pData = VarPtr(s) + ptrSz
    lpRef(0) = VarPtr(SA)
End Sub
Private Sub Test_VariantUnion()
    Dim vs, vbl, vlp, vd
    Initialize
    
    vs = "sdafasfdas"
    vbl = True
    vlp = VarPtr(vlp)
    vd = 3343.0809
    
'    vRef_SA.pData = VarPtr(vs)
'    Debug.Print vsRef(0).val
    vRef_SA.pData = VarPtr(vbl)
    Debug.Print viRef(0).val
    vRef_SA.pData = VarPtr(vlp)
    Debug.Print vlpRef(0).val
    vRef_SA.pData = VarPtr(vd)
    Debug.Print vdRef(0).val
End Sub
Private Sub Test_VariantUnion2()
    Dim pvArr As LongPtr, vArr(), _
        vsArr() As sVariant, vlpArr() As lpVariant, vdArr() As dVariant, viArr() As iVariant
    Dim i&, SA As SA1D, sRef$(), sRefSA As SA1D
    Initialize
    
    vArr = Array("строка", 344.887, "строка2", True, 323252)
    
    pvArr = VarPtr(pvArr) - ptrSz
    SA = GetSA(GetPtr(pvArr))
    SA.Features = FADF_AUTO Or FADF_FIXEDSIZE 'FADF_FIXEDSIZE_AUTO 'FADF_STATIC Or FADF_FIXEDSIZE
    SA.Locks = 1
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
            sRefSA.pData = VarPtr(vlpArr(i).val)
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
    Dim s1$, ref$(), SA As SA1D, emp$()
    Initialize
    
    s1 = "asfddafa"
    ref = RefStr(SA, VarPtr(s1))
'    Erase ref
    Debug.Print ref(0)
End Sub
'Function InStrLenRevNS(sCheck$, sMatch$, ByVal Start As Long, _
'    Optional ByVal Compare As VbCompareMethod, Optional ByVal lenFind As Long = -1) As Long
'    Dim sTmp$, pTmp As LongPtr, lTmp&, lOff&, lRet&
'  #If Not PreInitMode Then
'    If IsInitialized Then Else Initialize
'  #End If
'
'    lOff = Start - lenFind
'    pTmp = StrPtr(sCheck) + lOff * 2
'    lRef_SA.pData = pTmp - 4
'    lTmp = lRef(0)
'    lRef(0) = lenFind * 2
'    lpRef_SA.pData = VarPtr(sTmp)
'    lpRef(0) = pTmp
'
'    lRet = InStrRev(sTmp, sMatch, lenFind, Compare)
'    if lRet then InStrLenRevNS = lRet + lOff
'    lpRef(0) = 0
'    lRef(0) = lTmp
'End Function
