Attribute VB_Name = "DictionaryHelp"
Option Explicit
'https://www.cyberforum.ru/visual-basic/thread1146688.html
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VarBstrCmp Lib "oleaut32" (ByVal bstrLeft As LongPtr, ByVal bstrRight As LongPtr, ByVal lcid As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Sub VariantCopy Lib "oleaut32.dll" (pvargDest As Any, pvargSrc As Any)
 
Private Type tLong
    l As Long
End Type
Private Type tCurrency
    c As Currency
End Type
#If Win64 Then
    Private Const ptrSz = 8
    Private Const varSz = 24
    Private Const dictpFirstOffset = 48
    Private Const dictpHTblOffset = 64
    Private Const dictDivOffset = 72
    Private Const dictlcidOffset = 80
    Private Const dictSKeyOffset = 24
    
    Private Const dictVItemOffset = 40
    Private Const dictItemSize = 72
#Else
    Private Const ptrSz = 4
    Private Const varSz = 16
    Private Const dictpFirstOffset = 28
    Private Const dictpHTblOffset = 36
    Private Const dictDivOffset = 40
    Private Const dictlcidOffset = 48
    
    Private Const dictSKeyOffset = 16
    Private Const dictVItemOffset = 24
    Private Const dictItemSize = 40
#End If
Private Type tDictDescr
    lp1(4)  As LongPtr
    l1      As Long
    lCnt    As Long             '24
    pFirst  As LongPtr          '28
    lp2     As LongPtr
    pHTbl   As LongPtr          '36
    ldiv    As Long             '40
    lCmp    As VbCompareMethod  '44
    lcid    As Long             '48
End Type
Private Type tDictItem
    pInterface As LongPtr     '0
    pNext As LongPtr          '4
    Key As Variant            '8
    Item As Variant           '18 24
    PointerToHash As LongPtr  '28 40
    Reserved As Long          '2C 44
End Type

Private DictItemRef() As tDictItem, DictItemRef_SA As SA1D
Private DictItemRef2() As tDictItem, DictItemRef2_SA As SA1D
Private DictDescRef() As tDictDescr, DictDescRef_SA As SA1D
Private isDictHlpRefInit As Boolean

Private Sub TestDescr()
    Dim dict As New Dictionary, i&, pNext As LongPtr
    InitDictHlp
    dict.Add "key1", "item1"
    dict.Add "key2", "item2"
    dict.Add "key3", "item3"
    
    DictDescRef_SA.pData = ObjPtr(dict)
    With DictDescRef(0)
      DictItemRef_SA.pData = .pFirst
      Debug.Print .pFirst
      For i = 2 To .lCnt
          pNext = DictItemRef(0).pNext
          DictItemRef_SA.pData = pNext
          Debug.Print pNext
      Next
    End With
    
End Sub

Private Sub Example()
    Dim dict As New Dictionary
    
    dict.Add "key1", "item1"
    dict.Add "key2", "item2"
    dict.Add "key3", "item3"
'    Debug.Print DictItem(dict, "key2")
'    Debug.Print DictItemByIndex(dict, 2)
'    Debug.Print DictKeyByIndex(dict, 3)
'    Debug.Print Join(DictItems(dict), vbCr)
    DictRemoveByIndex dict, 2
End Sub

Private Sub InitDictHlp()
    Dim tdiTmp As tDictItem, tddTmp As tDictDescr
    If isDictHlpRefInit Then Exit Sub
    If IsInitialized Then Else Initialize
    
    MakeRef DictItemRef_SA, VarPtr(DictItemRef_SA) - ptrSz, LenB(tdiTmp)
    MakeRef DictItemRef2_SA, VarPtr(DictItemRef2_SA) - ptrSz, LenB(tdiTmp)
    MakeRef DictDescRef_SA, VarPtr(DictDescRef_SA) - ptrSz, LenB(tddTmp)
    
    isDictHlpRefInit = True
End Sub
' Получить элемент по индексу
Function DictItemByIndex(Dic As Dictionary, ByVal Index As Long) As Variant
    Dim pDesc As LongPtr, i&
    If isDictHlpRefInit Then Else InitDictHlp
    pDesc = ObjPtr(Dic)
    If pDesc Then Else GoTo errArgum
    DictDescRef_SA.pData = pDesc
    With DictDescRef(0)
      Select Case Index
      Case 1 To .lCnt
      Case Else: GoTo errArgum
      End Select
      
      DictItemRef_SA.pData = .pFirst
    End With
    For i = 2 To Index
        DictItemRef_SA.pData = DictItemRef(0).pNext
    Next
    
    DictItemByIndex = DictItemRef(0).Item
    DictItemRef_SA.pData = NullPtr
Exit Function
errArgum:
    DictItemRef_SA.pData = NullPtr
    Err.Raise 5, , "Bed argument value!"
End Function
' Получить ключ по индексу
Function DictKeyByIndex(Dic As Dictionary, ByVal Index As Long) As Variant
    Dim pDesc As LongPtr, i&
    If isDictHlpRefInit Then Else InitDictHlp
    pDesc = ObjPtr(Dic)
    If pDesc Then Else GoTo errArgum
    DictDescRef_SA.pData = pDesc
    With DictDescRef(0)
      Select Case Index
      Case 1 To .lCnt
      Case Else: GoTo errArgum
      End Select
      
      DictItemRef_SA.pData = .pFirst
    End With
    For i = 2 To Index
        DictItemRef_SA.pData = DictItemRef(0).pNext
    Next
    
    DictKeyByIndex = DictItemRef(0).Key
    DictItemRef_SA.pData = NullPtr
Exit Function
errArgum:
    DictItemRef_SA.pData = NullPtr
    Err.Raise 5, , "Bed argument value!"
End Function
' Удалить элемент по индексу
Sub DictRemoveByIndex(Dic As Dictionary, ByVal Index As Long)
    Dim pDesc As LongPtr, i&
    If isDictHlpRefInit Then Else InitDictHlp
    pDesc = ObjPtr(Dic)
    If pDesc Then Else GoTo errArgum
    DictDescRef_SA.pData = pDesc
    With DictDescRef(0)
      Select Case Index
      Case 1 To .lCnt
      Case Else: GoTo errArgum
      End Select
      
      DictItemRef_SA.pData = .pFirst
    End With
    For i = 2 To Index
        DictItemRef_SA.pData = DictItemRef(0).pNext
    Next
    
    Dic.Remove DictItemRef(0).Key
    DictItemRef_SA.pData = NullPtr
Exit Sub
errArgum:
    DictItemRef_SA.pData = NullPtr
    Err.Raise 5, , "Bed argument value!"
End Sub

' Получить элемент по ключу. Реплика функции The trick-а https://www.cyberforum.ru/visual-basic/thread1146688.html#post6040814
Private Function DictItem(Dic As Dictionary, Key As String) As Variant
    Dim Hash As Long, pHTbl As LongPtr, pHItem As LongPtr, pKey As LongPtr, cmp As Long ', lcid As Long
    If isDictHlpRefInit Then Else InitDictHlp
    DictDescRef_SA.pData = ObjPtr(Dic)
    cmp = DictDescRef(0).lCmp                         'Dic.CompareMode
'    lcid = DictDescRef(0).lcid                       ' Получаем lcid
    Hash = HashValVBA(DictDescRef(0), Key)            ' Вычисляем хэш
    pHTbl = DictDescRef(0).pHTbl                      ' Получаем указатель на хэш-таблицу
    pHItem = GetPtr(pHTbl + Hash * ptrSz)             ' Получаем указатель элемента в хэш-таблице
    Do While pHItem                                   ' Если есть такой элемент
        DictItemRef_SA.pData = pHItem
        If StrComp(Key, DictItemRef(0).Key, cmp) = 0 Then
            DictItem = DictItemRef(0).Item: Exit Function
        Else: pHItem = GetPtr(pHItem + dictItemSize)
        End If
    Loop
    DictItemRef_SA.pData = NullPtr
End Function
Private Function HashValVBA(DictDescr As tDictDescr, s As String) As Long
    Dim i&, ch@, lnStr&, sTmp$
    Dim res As tLong, cres As tCurrency
    With DictDescr
      If .lCmp = 0 Then
          iMap1_SA.pData = StrPtr(s)
      ElseIf .lCmp = TextCompare Then
          sTmp = LCase$(s)
          iMap1_SA.pData = StrPtr(sTmp)
      Else: iMap1_SA.pData = StrPtr(s)
      End If
      lnStr = Len(s)
      iMap1_SA.Count = lnStr
      ' Извращения с Currency, т.к. в VB нет UINT32 и циклической арифметики
      For i = 1 To lnStr
          ch = iMap1(i) / 10000
          cres.c = CCur(res.l) / 10000 * 17 + ch
          LSet res = cres
      Next
      cres.c = 0: LSet cres = res
      ch = cres.c * 10000
      HashValVBA = ch - (Int(ch / .ldiv) * .ldiv)
    End With
End Function
' Получить список элементов
Private Function DictItems(Dic As Dictionary) As Variant()
    Dim pItem As Long, vArOut() As Variant, i As Long, Ub&
    If isDictHlpRefInit Then Else InitDictHlp
    
    DictDescRef_SA.pData = ObjPtr(Dic)
    With DictDescRef(0)
      Ub = .lCnt - 1
      ReDim vArOut(Ub)
      DictItemRef_SA.pData = .pFirst
    End With
      
    vArOut(0) = DictItemRef(0).Item
    For i = 1 To Ub
        DictItemRef_SA.pData = DictItemRef(0).pNext
        vArOut(i) = DictItemRef(0).Item
    Next
    DictItemRef_SA.pData = NullPtr
    
    DictItems = vArOut
End Function
Function DictJoinedKeys(dict As Dictionary, Optional Dlm$ = " ") As String
    Dim i&, sRes$, resLen&, dlmLen&, newLen&, maxLen&, keyLen&
    Dim pRes As LongPtr, pDst As LongPtr, stpInc&, sTmp$, pTmp As LongPtr
    If isDictHlpRefInit Then Else InitDictHlp
    
    dlmLen = LenB(Dlm)
    DictDescRef_SA.pData = ObjPtr(dict)
    With DictDescRef(0)
      If .lCnt Then Else GoTo endFn
      stpInc = 8
      DictItemRef_SA.pData = .pFirst
      sRes = DictItemRef(0).Key ': Debug.Print StrPtr(sRes)
      pDst = StrPtr(sRes)
      resLen = LenB(sRes)
      pTmp = VarPtr(sTmp)
      For i = 2 To DictDescRef(0).lCnt
          DictItemRef_SA.pData = DictItemRef(0).pNext
          With DictItemRef(0)
            If VarType(.Key) = vbString Then
                keyLen = LenB(.Key)
                sRef2_SA.pData = VarPtr(.Key) + 8
            Else
                sTmp = .Key
                keyLen = LenB(sTmp)
                sRef2_SA.pData = pTmp
            End If
            newLen = resLen + dlmLen + keyLen
            If newLen > maxLen Then
                Do
                    maxLen = maxLen + stpInc
                    stpInc = stpInc * 2
                Loop While newLen > maxLen
                ReallocStringB sRes, maxLen ': Debug.Print StrPtr(sRes)
                pRes = StrPtr(sRes)
            End If
            pDst = pRes + resLen
            PutStrBuf pDst, Dlm
            pDst = pDst + dlmLen
            PutStrBuf pDst, sRef2(0)
            
            resLen = newLen
          End With
      Next
    End With
    ReallocStringB sRes, resLen ': Debug.Print StrPtr(sRes)
    
    MoveStr DictJoinedKeys, sRes
endFn:
    DictItemRef_SA.pData = 0
    DictDescRef_SA.pData = 0
End Function

Private Sub Test_DictJoinedKeys()
    Dim dict As New Dictionary
    Dim s$
    dict.Add "key1", "item1"
    dict.Add 5, "item2"
    dict.Add 57.33, "item3"
    dict.Add "key4", "item4"

    s = DictJoinedKeys(dict)
End Sub
'' Получить список элементов
'Private Function Items(Dic As Dictionary) As Variant
'    Dim pItem As Long, loc() As Variant, i As Long
'
'    ReDim loc(Dic.Count - 1)
'
'    GetMem4 ByVal ObjPtr(Dic) + dictpFirstOffset, pItem ' Указатель на первый элемент списка
'    Do                                        ' Проход по элементам списка
'        VariantCopy loc(i), ByVal pItem + &H18
'        GetMem4 ByVal pItem + 4, pItem          ' Следующий элемент
'        i = i + 1
'    Loop While pItem
'
'    Items = loc
'End Function
'' Получить элемент по индексу
'Private Function DictItemByIndex(Dic As Dictionary, ByVal Index As Long) As Variant
'    Dim pItem As Long
'    If Dic.Count = 0 Then Exit Function
'
'    GetMem4 ByVal ObjPtr(Dic) + &H1C, pItem ' Указатель на первый элемент списка
'
'    Do While CBool(Index) And pItem ' Проход по элементам списка
'        GetMem4 ByVal pItem + 4, pItem ' Следующий элемент
'        Index = Index - 1
'    Loop
'
'    VariantCopy DictItemByIndex, ByVal pItem + &H18
'End Function

'' Получить список элементов
'Private Function Items(Dic As Dictionary) As Variant
'    Dim pItem As Long, loc() As Variant, i As Long
'
'    ReDim loc(Dic.Count - 1)
'
'    GetMem4 ByVal ObjPtr(Dic) + dictpFirstOffset, pItem ' Указатель на первый элемент списка
'    Do                                        ' Проход по элементам списка
'        VariantCopy loc(i), ByVal pItem + &H18
'        GetMem4 ByVal pItem + 4, pItem          ' Следующий элемент
'        i = i + 1
'    Loop While pItem
'
'    Items = loc
'End Function
'Private Function GetItemVBA(Dic As Dictionary, Key As String) As Variant
'    Dim Hash As Long, pHTbl As LongPtr, pHItem As LongPtr, lcid As Long, pKey As LongPtr, cmp As Long
'    Dim pDic As LongPtr
'    pDic = ObjPtr(Dic)
'    cmp = Dic.CompareMode
'    CopyMemory lcid, ByVal pDic + dictlcidOffset, ptrSz             ' Получаем lcid
'    Hash = HashValVBA(Dic, Key)                                     ' Вычисляем хэш
'    CopyMemory pHTbl, ByVal pDic + dictpHTblOffset, ptrSz           ' Получаем указатель на хэш-таблицу
'
'    CopyMemory pHItem, ByVal pHTbl + Hash * ptrSz, ptrSz            ' Получаем указатель элемента в хэш-таблице
'    Do While pHItem                                                 ' Если есть такой элемент
'        CopyMemory pKey, ByVal pHItem + dictSKeyOffset, ptrSz       ' Сравниваем значение ключа в таблице с заданым ключем
'        Select Case VarBstrCmp(StrPtr(Key), pKey, lcid, cmp)
'        Case 1: VariantCopy GetItemVBA, ByVal pHItem + dictVItemOffset: Exit Function
'        Case Else
'            CopyMemory pHItem, ByVal pHItem + dictItemSize, ptrSz  ' Получаем указатель на следующую запись в таблице
'        End Select
'    Loop
'End Function

