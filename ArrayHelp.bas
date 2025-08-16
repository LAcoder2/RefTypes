Attribute VB_Name = "ArrayHelp"
Option Explicit

Public Type SABounds
    Count As Long
    lBound As Long
End Type
Public Type SA2D '(SAFEARRAY2D)
    Dims          As Integer
    Features      As Integer
    cbElem        As Long
    Locks         As Long
  #If Win64 Then
    padding       As Long
  #End If
    pData         As LongPtr
    ColCount      As Long
    CollBound     As Long
    RowCount      As Long
    RowlBound     As Long
End Type
Public Enum Linear2DArrayType
    RowVector
    ColumnVector
End Enum
Public sa2dRef() As SA2D, sa2dRef_SA As SA1D
Private isArrHlpInit As Boolean

Private Sub InitArrHlp()
    Dim SA2dTmp As SA2D
    If isArrHlpInit Then Exit Sub
    If IsInitialized Then Else Initialize
    
    MakeRef sa2dRef_SA, VarPtr(sa2dRef_SA) - ptrSz, LenB(SA2dTmp) 'ссылка на структуру SafeArray2D
    
    isArrHlpInit = True
End Sub
'Join вариантного 2D массива
'Принцип работы: временно преобразует 2d-массив в 1d, выполняет Join() его содержимого и вновь преобразует из 2d в 1d.
Function JoinV2D(vAr2D(), Optional Delim$ = " ") As String
    Dim ColCount&
    If isArrHlpInit Then Else InitArrHlp
    sa2dRef_SA.pData = ArrPtrV(vAr2D, True)
    With sa2dRef(0)
      If .Dims = 2 Then Else GoTo errArgum
      .Dims = 1
      ColCount = .ColCount
      .ColCount = .ColCount * .RowCount
      JoinV2D = Join(vAr2D, Delim)
      .ColCount = ColCount
      .Dims = 2
    End With
Exit Function
errArgum:
    Err.Raise 5, , "Arguments error!"
End Function
Sub vAry2Dto1D(vAry())
    If isArrHlpInit Then Else InitArrHlp
    sa2dRef_SA.pData = ArrPtrV(vAry, True)
    With sa2dRef(0)
      If .Dims = 2 Then Else GoTo errArgum
      .Dims = 1
      .ColCount = .ColCount * .RowCount
    End With

Exit Sub
errArgum:
    Err.Raise 5, , "Arguments error!"
End Sub
Sub vAry1Dto2D(vAry(), AryType As Linear2DArrayType)
    Dim pvAry As LongPtr, pSA As LongPtr, SA2dTmp As SA2D
    If isArrHlpInit Then Else InitArrHlp
    pvAry = ArrPtrV(vAry)
    pSA = VbaMemRealloc(GetPtr(pvAry), LenB(SA2dTmp))
    PutPtr(pvAry) = pSA
    sa2dRef_SA.pData = pSA
    With sa2dRef(0)
        If .Dims = 1 Then Else GoTo errArgum
        If AryType = ColumnVector Then
            .RowCount = .ColCount
            .ColCount = 1
        Else
            .RowCount = 1
        End If
        .Dims = 2
    End With

Exit Sub
errArgum:
    Err.Raise 5, , "Arguments error!"
End Sub
Private Sub Test_RedimPreserve2DVectorV()
    Dim vArr()
    ReDim vArr(1 To 5, 1 To 1)
    
    RedimPreserve2DColumnVectorV vArr, 10
End Sub
Sub RedimPreserve2DColumnVectorV(vAry(), ByVal newBound As LongPtr)
    Dim colCnt&, collBnd&
    If isArrHlpInit Then Else InitArrHlp
    
    sa2dRef_SA.pData = ArrPtrV(vAry, True)
    With sa2dRef(0)
      If .Dims = 2 Then Else GoTo errArgum  'isn't 2d
      .Dims = 1
      If .ColCount = 1 Then Else: GoTo errArgum 'isn't ColumnVector
      colCnt = .ColCount
      collBnd = .CollBound
      .ColCount = .RowCount
      .CollBound = .RowlBound
      ReDim Preserve vAry(.RowlBound To newBound)
      .RowCount = newBound
      .ColCount = colCnt
      .CollBound = collBnd
      .Dims = 2
    End With
Exit Sub
errArgum:
    Err.Raise 5, , "Arguments error!"
End Sub
Sub Test_vAry1Dto2D_2Dto1D()
    Dim vAry()
    
    ReDim vAry(4, 3)
    Debug.Print ArrPtrV(vAry(), True)
    vAry2Dto1D vAry
    Debug.Print ArrPtrV(vAry(), True)
    vAry1Dto2D vAry, ColumnVector
    Debug.Print ArrPtrV(vAry(), True)
End Sub

Function SAAllocDescr(ByVal Dims As Integer) As LongPtr
    Const szBnds& = 8
    Static init As Boolean, szDesc1D&
    Dim sTmp$
    If init Then
    Else
        If isArrHlpInit Then Else InitArrHlp
        szDesc1D = LenB(saRef_SA)
    End If
    If Dims > 0 Then Else Exit Function
    
'    MovePtr VarPtr(SAAllocDescr), VarPtr(String((szDesc1D + (Dims - 1) * szBnds) \ 2 - 3, vbNullChar)) + 8
'    SAAllocDescr = SAAllocDescr - 4
    SAAllocDescr = VbaMemAlloc(szDesc1D + (Dims - 1) * szBnds)
    saRef_SA.pData = SAAllocDescr
    With saRef(0)
        .Dims = Dims
        .Features = &H80
    End With
End Function

Private Sub Test_SAAllocDescr()
    Dim pDesc As LongPtr
    pDesc = SAAllocDescr(2)
    VbaMemFree pDesc
End Sub
Private Sub Test_JoinV2D()
    Dim arr(), sRes$
    
    arr = Selection.Value
    sRes = JoinV2D(arr, vbLf)
    Debug.Print sRes
End Sub

Private Sub TestSA2D()
    Dim arr(), pArr As LongPtr, pSA As LongPtr, SA As SA2D, sTmp$
    ReDim arr(1 To 5, 1 To 2)
    InitArrHlp
    sTmp = JoinV2D(arr)
    
    pArr = VarPtr(pArr) + ptrSz
    pSA = GetPtr(pArr)
    
'    MemLSet VarPtr(SA), pSA, LenB(SA)
    sa2dRef_SA.pData = pSA
    
End Sub
