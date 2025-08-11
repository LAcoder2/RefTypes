Attribute VB_Name = "CollectionHelp"
Option Explicit
'Functions to extend the functionality of VB/VBA collections.
'Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
'Private Declare PtrSafe Function ArrPtr Lib "vbe7" Alias "VarPtr" (Arr() As Any) As LongPtr
'Private Type pVariant
'    cur As Currency
'    ptr As LongPtr
'    lp0 As LongPtr
'End Type
'Private Type tpCollElemPtr
'    pItem As pVariant
'    pKey As LongPtr
'    prvPtr As LongPtr
'    nxtPtr As LongPtr
'End Type
'Private Type tpCollection
'    pInterface1         As IUnknown            ' // 0x00
'    pInterface2         As IUnknown            ' // 0x04
'    pInterface3         As IUnknown            ' // 0x08
'    lRefCounter         As Long                ' // 0x0C
'    lNumOfItems         As Long                ' // 0x10
'    pvUnk1              As LongPtr             ' // 0x14
'    pFirstIndexedItem   As LongPtr             ' // 0x18
'    pLastIndexedItem    As LongPtr             ' // 0x1C
'    pvUnk4              As LongPtr             ' // 0x20
'    pFirstItem          As LongPtr             ' // 0x24
'    pRootItem           As LongPtr             ' // 0x28
'    pvUnk5              As LongPtr             ' // 0x2C
'End Type
Private Type tCollItem
    vItem         As Variant    '0  0   0
    sKey          As String     '16 10  24
    pPrev         As LongPtr    '20 14  32
    pNext         As LongPtr    '24 18  40
    pUnknown      As LongPtr    '28 1C  48
    pParent       As LongPtr    '32 20  56
    pRight        As LongPtr    '36 24  64
    pLeft         As LongPtr    '40 28  72
    bFlag         As Boolean    '44 2C  80
End Type

Private Type GatherKeysInOrderStack
    Count As Long
    Offset1 As Long
    offset2 As Long
    pRoot As LongPtr
End Type

Private CollItemRef() As tCollItem, CollItemRef_SA As SA1D, CollItemRef2() As tCollItem, CollItemRef2_SA As SA1D
'Private tCollRef() As tpCollection, tCollRef_SA As SA1D
Private isCollItemRefInit As Boolean
'#If Win64 Then
'    Private Const RightOffset = 64
'    Private Const LeftOffset = 72
''    Private Const collItemOffset = 40
'#Else
    Private Const RightOffset = varSz + ptrSz * 5   '36
    Private Const LeftOffset = RightOffset + ptrSz '40
'    Private Const collItemOffset = 24
'#End If
Const NullPtr As LongPtr = 0

Private Sub Example()
    Dim coll As New VBA.Collection
    
    coll.Add "Строка 1", "Дерево"
    coll.Add "Строка 2", "Арбуз"
    coll.Add "Строка 3", "Банан"
    coll.Add "Строка 4" ', "Аппельсин"
    coll.Add "Строка 5", "Ананас"
    coll.Add "Строка 6", "груша"
    coll.Add "Строка 7" ', "вишня"
    coll.Add "Строка 8", "абрикос"
    
    Debug.Print CollItem("Арбуз", coll)
    Debug.Print CollExists("вишня", coll)
    Debug.Print CollExists("абрикос", coll)
    Debug.Print CollKeyByIndex(2, coll)
    Debug.Print
    
    Dim keys$(), Items()
    keys = CollKeys(coll)
    Items = CollItems(coll)
    Debug.Print Join2(keys) & vbCr & _
                Join(Items)
    Debug.Print
    
    keys = CollSortedKeys(coll)
    Debug.Print Join2(keys, vbCr)
    Debug.Print
    Debug.Print CollKeyIndex("Ананас", coll)
'    Debug.Print coll("Банан")
    Debug.Print CollRemoveByKey("Банан", coll)
'    Debug.Print coll("Банан")
    
End Sub

Private Sub InitCollItemRef()
    Dim tciTmp As tCollItem ', tcTmp As tpCollection
    If isCollItemRefInit Then Exit Sub
    If IsInitialized Then Else Initialize
    
    MakeRef CollItemRef_SA, VarPtr(CollItemRef_SA) - ptrSz, LenB(tciTmp)
    MakeRef CollItemRef2_SA, VarPtr(CollItemRef2_SA) - ptrSz, LenB(tciTmp)
'    MakeRef tCollRef_SA, VarPtr(tCollRef_SA) - ptrSz, LenB(tcTmp)
    
    isCollItemRefInit = True
End Sub

'https://www.cyberforum.ru/visual-basic/thread1096760.html
'https://www.cyberforum.ru/visual-basic/thread1801288.html
Private Function CollKeyByIndex(ByVal Index As Long, coll As Collection) As String
    Dim i As Long
    If coll Is Nothing Then Exit Function
    If isCollItemRefInit Then Else InitCollItemRef
    
    Select Case Index
    Case 1 To coll.Count
        CollItemRef_SA.pData = ObjPtr(coll)
        For i = 1 To Index
            CollItemRef_SA.pData = CollItemRef(0).pNext
        Next
    Case Else: Exit Function
    End Select
    
    CollKeyByIndex = CollItemRef(0).sKey
    CollItemRef_SA.pData = 0
End Function
Function CollKeys(coll As VBA.Collection) As String()
    Dim keys$(), i&, Ub&
    If isCollItemRefInit Then Else InitCollItemRef
    
    Ub = coll.Count - 1
    If Ub > -1 Then Else Exit Function
    ReDim keys(Ub)
    CollItemRef_SA.pData = ObjPtr(coll)
    For i = 0 To Ub
        CollItemRef_SA.pData = CollItemRef(0).pNext
        keys(i) = CollItemRef(0).sKey
    Next
    CollItemRef_SA.pData = 0
    
    CollKeys = keys
End Function
Function CollItems(coll As VBA.Collection) As Variant()
    Dim Items(), i&, Ub&, Key$
    If isCollItemRefInit Then Else InitCollItemRef
    
    Ub = coll.Count - 1
    If Ub = -1 Then Exit Function
    ReDim Items(Ub)
    CollItemRef_SA.pData = ObjPtr(coll)
    For i = 0 To Ub
        CollItemRef_SA.pData = CollItemRef(0).pNext
        Items(i) = CollItemRef(0).vItem
    Next
    CollItemRef_SA.pData = 0
    
    CollItems = Items
End Function

Function CollItem(Key As String, Col As VBA.Collection) As Variant
    Dim pItem As Long, pRoot As Long
    If isCollItemRefInit Then Else InitCollItemRef
    
    CollItemRef_SA.pData = ObjPtr(Col)
    pRoot = CollItemRef(0).pLeft            'pRootItem
    pItem = CollItemRef(0).pRight           'pFirstItem
    CollItemRef_SA.pData = pItem
    
    Do Until pItem = pRoot
        Select Case StrComp(Key, CollItemRef(0).sKey) ' , vbTextCompare)??
        Case -1: pItem = CollItemRef(0).pLeft     'если меньше
        Case 0                                    'если равны
            CollItem = CollItemRef(0).vItem
            CollItemRef_SA.pData = 0
            Exit Function
        Case Else: pItem = CollItemRef(0).pRight  'если больше
        End Select
        CollItemRef_SA.pData = pItem
    Loop
    CollItemRef_SA.pData = 0
    
    MsgBox "Element not found"
End Function
Function CollExists(Key As String, Col As VBA.Collection) As Boolean
    Dim pItem As Long, pRoot As Long
    If isCollItemRefInit Then Else InitCollItemRef
    
    CollItemRef_SA.pData = ObjPtr(Col)
    pRoot = CollItemRef(0).pLeft            'pRootItem
    pItem = CollItemRef(0).pRight           'pFirstItem
    CollItemRef_SA.pData = pItem
    
    Do Until pItem = pRoot
        Select Case StrComp(Key, CollItemRef(0).sKey)
        Case -1: pItem = CollItemRef(0).pLeft   'если меньше
        Case 0                                  'если равны
            CollItemRef_SA.pData = 0
            CollExists = True: Exit Function
        Case Else: pItem = CollItemRef(0).pRight 'если больше
        End Select
        CollItemRef_SA.pData = pItem
    Loop
    CollItemRef_SA.pData = 0
End Function
Function CollRemoveByKey(Key As String, Col As VBA.Collection) As Boolean
    Dim pItem As Long, pRoot As Long, lIndex&
    If isCollItemRefInit Then Else InitCollItemRef
    
    CollItemRef_SA.pData = ObjPtr(Col)
    pRoot = CollItemRef(0).pLeft            'pRootItem
    pItem = CollItemRef(0).pRight           'pFirstItem
    CollItemRef_SA.pData = pItem
    
    Do Until pItem = pRoot
        Select Case StrComp(Key, CollItemRef(0).sKey) ' , vbTextCompare)??
        Case -1: pItem = CollItemRef(0).pLeft     'если меньше
        Case 0                                    'если равны
            lIndex = 1
            Do
                With CollItemRef(0)
                  If .pPrev Then
                      lIndex = lIndex + 1
                      CollItemRef_SA.pData = .pPrev
                  Else: Exit Do
                  End If
                End With
            Loop
            Col.Remove lIndex
            CollItemRef_SA.pData = 0
            CollRemoveByKey = True: Exit Function 'RETURN
        Case Else: pItem = CollItemRef(0).pRight  'если больше
        End Select
        CollItemRef_SA.pData = pItem
    Loop
    CollItemRef_SA.pData = 0
End Function
Function CollKeyIndex(Key As String, Col As VBA.Collection) As Long
    Dim pItem As Long, pRoot As Long, lIndex&
    If isCollItemRefInit Then Else InitCollItemRef
    
    CollItemRef_SA.pData = ObjPtr(Col)
    pRoot = CollItemRef(0).pLeft            'pRootItem
    pItem = CollItemRef(0).pRight           'pFirstItem
    CollItemRef_SA.pData = pItem
    
    Do Until pItem = pRoot
        Select Case StrComp(Key, CollItemRef(0).sKey) ' , vbTextCompare)??
        Case -1: pItem = CollItemRef(0).pLeft     'если меньше
        Case 0                                    'если равны
            lIndex = 1
            Do
                With CollItemRef(0)
                  If .pPrev Then
                      lIndex = lIndex + 1
                      CollItemRef_SA.pData = .pPrev
                  Else: Exit Do
                  End If
                End With
            Loop
            CollKeyIndex = lIndex 'RETURN
            CollItemRef_SA.pData = 0
            Exit Function
        Case Else: pItem = CollItemRef(0).pRight  'если больше
        End Select
        CollItemRef_SA.pData = pItem
    Loop
    CollItemRef_SA.pData = 0
    
    MsgBox "Element not found"
End Function

'https://www.vbforums.com/showthread.php?868451-Iterate-thru-VB6-Collection-in-Alphabetic-Key-Order
Function CollSortedKeys(coll As Collection, Optional ByVal blReverse As Boolean) As String()
    ' Originally written by Wqweto, tweaked by Elroy.
    ' Returns 0 to -1 array on empty Collection.
    ' This is particularly nice when you want to use the Collection for nothing but sorting.
    ' Does NOT return items with no key.    '
    Dim pFirst As Long, lCnt&, argStack As GatherKeysInOrderStack
    Select Case True
    Case coll Is Nothing, coll.Count = 0: Exit Function
    End Select
    If isCollItemRefInit Then Else InitCollItemRef
    
    CollItemRef_SA.pData = ObjPtr(coll)
    pFirst = CollItemRef(0).pRight
    With argStack
      .pRoot = CollItemRef(0).pLeft
      
      If pFirst = .pRoot Then CollItemRef_SA.pData = 0: Exit Function
      
      ' Offsets that determine forward or reverse.
      If Not blReverse Then
          .Offset1 = LeftOffset      ' pLeftBranch
          .offset2 = RightOffset     ' pRightBranch
      Else
          .Offset1 = RightOffset     ' pRightBranch
          .offset2 = LeftOffset      ' pLeftBranch
      End If
      
      ' Gather the keys.
      ReDim CollSortedKeys(1 To coll.Count)
      GatherKeysInOrder pFirst, CollSortedKeys, argStack
      If .Count < coll.Count Then ReDim Preserve CollSortedKeys(1 To .Count)
    End With
    CollItemRef_SA.pData = 0
End Function
Private Sub GatherKeysInOrder(ByVal pItem As LongPtr, sKeys() As String, argStack As GatherKeysInOrderStack)
    ' Originally written by Wqweto, tweaked by Elroy and Testuser2(2025)
    Dim pNewItem As Long
    
    With argStack
        pNewItem = GetPtr(pItem + .Offset1)        ' Traverse left (or right, if reverse) branch if present.
        If pNewItem <> .pRoot Then GatherKeysInOrder pNewItem, sKeys, argStack
        
        .Count = .Count + 1
        CollItemRef_SA.pData = pItem
        sKeys(.Count) = CollItemRef(0).sKey
        
        ' Traverse right (or left, if reverse) branch if present.
        pNewItem = GetPtr(pItem + .offset2)
        If pNewItem <> .pRoot Then GatherKeysInOrder pNewItem, sKeys, argStack
    End With
End Sub


'Function CollKeys(coll As VBA.Collection) As String()
'    Dim i&, pItem As LongPtr, Key$, pKey As LongPtr
'    Dim Ub&: Ub = coll.Count - 1
'    If Ub = -1 Then Exit Function
'    Dim keys$(): ReDim keys(Ub)
'    pKey = VarPtr(Key)
'    pItem = ObjPtr(coll)
'    For i = 0 To Ub
''        CopyMemory pItem, ByVal pItem + collItemOffset, ptrSz
''        CopyMemory ByVal pKey, ByVal pItem + varSz, ptrSz
'        GetMem4 ByVal pItem + collItemOffset, pItem
'        GetMem4 ByVal pItem + varSz, ByVal pKey
'        keys(i) = Key
'    Next
'    CopyMemory ByVal VarPtr(Key), NullPtr, ptrSz
'    CollKeys = keys
'End Function
'Function CollItems(coll As VBA.Collection) As Variant()
'    Dim i&, pItem As LongPtr, item
'    Dim Ub&: Ub = coll.Count - 1
'    If Ub = -1 Then Exit Function
'    Dim Items(): ReDim Items(Ub)
'    pItem = ObjPtr(coll)
'    For i = 0 To Ub
''        CopyMemory pItem, ByVal pItem + collItemOffset, ptrSz
'        GetMem4 ByVal pItem + collItemOffset, pItem
'        CopyMemory item, ByVal pItem, varSz
'        If IsObject(item) Then
'            Set Items(i) = item
'        Else: Items(i) = item
'        End If
'    Next
'    CopyMemory item, Empty, varSz
'    CollItems = Items()
'End Function
'Private Sub CollTest1()
'    Dim i&, coll As New VBA.Collection, Elem
'    Dim ptr As LongPtr, item, Key$, empVar, Ptr0 As LongPtr
'
'    coll.Add "item1", "key1"
'    coll.Add "item2", "key2"
'    coll.Add "item3", "key3"
'
'    ptr = ObjPtr(coll)
'    For i = 1 To coll.Count
'        CopyMemory ptr, ByVal ptr + collItemOffset, ptrSz
''        CopyMemory Item, ByVal Ptr, varSz
'        CopyMemory ByVal VarPtr(Key), ByVal ptr + varSz, ptrSz
'        Debug.Print Key, item
''        Debug.Print StrPtr(Item)
'    Next
''    For Each Elem In Coll
''        Debug.Print StrPtr(Elem)
''    Next
''    CopyMemory Item, ByVal VarPtr(empVar), varSz
'    CopyMemory ByVal VarPtr(Key), Ptr0, ptrSz
'End Sub
'Sub testTurboDictionary()
'    Dim dic As New TurboDictionary
'
'    dic.Add "1val", "latin"
'    dic.Add "значен", "кирилл"
'    Debug.Print dic.item("1val")
'    Debug.Print dic.item("значен")
'End Sub
'Получение массива строк-указателей ключей коллекции
'Private Sub CollKeys(Coll As Collection, Keys As sArray, pKeys() As LongPtr)
'    Dim collElem As tpCollElemPtr
'    CollKeys_ Coll, Keys, pKeys, collElem
'End Sub
'Private Sub CollKeys_(Coll As Collection, Keys As sArray, pKeys() As LongPtr, collElem As tpCollElemPtr, _
'                      Optional Ptr As LongPtr, Optional pPtr As LongPtr, Optional ByVal Ptr0 As LongPtr)
'    Dim i As Long, Cnt As Long, lpTmp As LongPtr, ptKeys As LongPtr
'
'    If Coll Is Nothing Then Exit Sub
'    Cnt = Coll.Count
'    If Cnt Then
'        ReDim Keys.sArr(1 To Cnt, 1 To 1)
'        CopyMemory ByVal VarPtr(Ptr0) - ptrSz, VarPtr(Ptr0) - ptrSz * 2, ptrSz
'        pPtr = VarPtr(Keys): lpTmp = Ptr
'        pPtr = ArrPtr(pKeys): Ptr = lpTmp
'        pPtr = ObjPtr(Coll) + collItemOffset: lpTmp = Ptr
'        pPtr = VarPtr(Ptr0) - ptrSz * 3: Ptr = lpTmp
'
'        pKeys(1, 1) = collElem.pKey
'        For i = 2 To Cnt
'            Ptr = collElem.nxtPtr
'            pKeys(i, 1) = collElem.pKey
'        Next
'    End If
'End Sub
'Получение массива значений/указателей коллекции
'Private Sub CollItems(Coll As Collection, Items() As Variant, pItems() As pVariant)
'    Dim collElem As tpCollElemPtr
'    CollItems_ Coll, Items, pItems, collElem
'End Sub
'Private Sub CollItems_(Coll As Collection, Items() As Variant, pItems() As pVariant, collElem As tpCollElemPtr, _
'                      Optional Ptr As LongPtr, Optional pPtr As LongPtr, Optional ByVal Ptr0 As LongPtr)
'    Dim i As Long, Cnt As Long, lpTmp As LongPtr, ptItems As LongPtr
'
'    If Coll Is Nothing Then Exit Sub
'    Cnt = Coll.Count
'    If Cnt Then
'        ReDim Items(1 To Cnt, 1 To 1)
'        CopyMemory ByVal VarPtr(Ptr0) - ptrSz, VarPtr(Ptr0) - ptrSz * 2, ptrSz
'        pPtr = ArrPtr(Items): lpTmp = Ptr
'        pPtr = ArrPtr(pItems): Ptr = lpTmp
'        pPtr = ObjPtr(Coll) + collItemOffset: lpTmp = Ptr
'        pPtr = VarPtr(Ptr0) - ptrSz * 3: Ptr = lpTmp
'
'        pItems(1, 1) = collElem.pItem
'        For i = 2 To Cnt
'            Ptr = collElem.nxtPtr
'            pItems(i, 1) = collElem.pItem
'        Next
'    End If
'End Sub
'Private Sub CollKeys2(Coll As Collection, Keys() As String, pKeys() As LongPtr)
'    Dim collElem As tpCollElemPtr
'    CollKeys2_ Coll, Keys, pKeys, collElem
'End Sub
'Private Sub CollKeys2_(Coll As Collection, Keys() As String, pKeys() As LongPtr, collElem As tpCollElemPtr, _
'                      Optional Ptr As LongPtr, Optional pPtr As LongPtr, Optional ByVal Ptr0 As LongPtr)
'    Dim i As Long, Cnt As Long, lpTmp As LongPtr, ptKeys As LongPtr
'
'    If Coll Is Nothing Then Exit Sub
'    Cnt = Coll.Count
'    If Cnt Then
'        ReDim Keys(1 To Cnt, 1 To 1)
'        CopyMemory ByVal VarPtr(Ptr0) - ptrSz, VarPtr(Ptr0) - ptrSz * 2, ptrSz
'        Ptr0 = VarPtr(Ptr0)
'        pPtr = Ptr0 - ptrSz * 5: pPtr = Ptr: lpTmp = Ptr 'получение указателя Keys()
'        pPtr = Ptr0 - ptrSz * 4: pPtr = Ptr: Ptr = lpTmp '2й вариант: pPtr = ArrPtr(pKeys): Ptr = lpTmp
'        pPtr = ObjPtr(Coll) + collItemOffset: lpTmp = Ptr
'        pPtr = Ptr0 - ptrSz * 3: Ptr = lpTmp
'
'        pKeys(1, 1) = collElem.pKey
'        For i = 2 To Cnt
'            Ptr = collElem.nxtPtr
'            pKeys(i, 1) = collElem.pKey
'        Next
'    End If
'End Sub
 
'Private Sub ПримерCollKeyByIndexs()
'    Dim Coll As New Collection
'    Dim Ptr0 As LongPtr, pKeys() As LongPtr, Keys() As String
'    Dim i&, Cnt&
'
'    Coll.Add "item1", "key1"
'    Coll.Add "item2", "key2"
'    Coll.Add "item3", "key3"
'    Cnt = Coll.Count
'
'    CollKeyByIndexs Coll, Keys, pKeys
'
'    Range("A1").Resize(Cnt).Value = Keys
'
'    'освобождение указателей
''    For i = 1 To Cnt
''        pKeys(i, 1) = 0
''    Next
'    Erase pKeys
'    CopyMemory ByVal VarPtr(Ptr0) - ptrSz, Ptr0, ptrSz 'CopyMemory ByVal ArrPtr(pKeys), Ptr0, ptrSz
'End Sub
'Private Function CollItem(ByVal Index As Long, coll As Collection) As Variant
'    Dim i As Long, Ptr As LongPtr, Ptr2 As LongPtr, item As Variant, empVar As Variant
'    If coll Is Nothing Then Exit Function
'    Select Case Index
'    Case Is < 1, Is > coll.Count: Exit Function
'    Case Else
'        Ptr = ObjPtr(coll)
'        For i = 1 To Index
'            CopyMemory Ptr, ByVal Ptr + collItemOffset, ptrSz
'        Next
'    End Select
'    CopyMemory item, ByVal Ptr, varSz
'    CollItem = item
'    CopyMemory item, ByVal VarPtr(empVar), varSz
'End Function
'Private Function CollKeyByIndex(ByVal Index As Long, coll As Collection) As String
'    Dim i As Long, Ptr0 As LongPtr, Ptr As LongPtr, Key As String
'    If coll Is Nothing Then Exit Function
'    Select Case Index
'    Case Is < 1, Is > coll.Count: Exit Function
'    Case Else
'        Ptr = ObjPtr(coll)
'        For i = 1 To Index
'            CopyMemory Ptr, ByVal Ptr + collItemOffset, ptrSz
'        Next
'    End Select
'    CopyMemory ByVal VarPtr(Key), ByVal Ptr + varSz, ptrSz
'    CollKeyByIndex = Key
'    CopyMemory ByVal VarPtr(Key), Ptr0, ptrSz
'End Function
