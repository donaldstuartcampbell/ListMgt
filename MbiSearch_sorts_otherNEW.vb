Module MbiSearch_sorts_otherNEW
    Structure stringTwoIntegerType
        Dim str As String
        Dim ptr1 As Integer
        Dim ptr2 As Integer
    End Structure
    Structure stringThreeIntegerType
        Dim str As String
        Dim ptr1 As Integer
        Dim ptr2 As Integer
        Dim ptr3 As Integer
    End Structure

    Structure DateStringType
        Dim xDate As DateTime
        Dim Str As String
    End Structure

    'Structure StringPointerType
    '    Dim str As String
    '    Dim ptr As Integer
    'End Structure
    Structure DateIntegerType
        Dim xDate As Date
        Dim I1 As Integer
    End Structure
    Structure PointerStringType
        Dim ptr As Integer
        Dim str As String
    End Structure
    Structure StringPointerLongType
        Dim str As String
        Dim ptr As Long
    End Structure

    Sub sort_StringTwoInteger(ByRef a() As stringTwoIntegerType)
        System.Array.Sort(a, (Function(x As stringTwoIntegerType, y As stringTwoIntegerType) String.Compare(x.str, y.str))) 'lambda function
    End Sub
    Sub sort_StringPointer(ByRef a() As StringPointerType)
        System.Array.Sort(a, (Function(x As StringPointerType, y As StringPointerType) String.Compare(x.str, y.str))) 'lambda function
    End Sub
    Sub sort_String(ByRef a() As String)
        System.Array.Sort(a)
    End Sub
    Sub test_BiSearch()
        Dim a(0) As Integer
        Dim i As Integer = 0
        setSearchArray_integer(a)

        'Debug.Print(BiSearch_Int_GTE(a, 54, , False))
        'Debug.Print(BiSearch_Int_GTE(a, 54)) 'same as above
        'Debug.Print(BiSearch_Int_GTE(a, 54, , True))

        'Debug.Print(BiSearch_Int_LTE(a, 30, , True))
        'Debug.Print(BiSearch_Int_LTE(a, 54, , True))
        'Debug.Print(BiSearch_Int_LTE(a, 55, , True))

        Debug.Print(BiSearch_Int_GT(a, 29).ToString)
        Debug.Print(BiSearch_Int_GT(a, 30).ToString)
        Debug.Print(BiSearch_Int_GT(a, 54).ToString)
        Debug.Print(BiSearch_Int_GT(a, 55).ToString)

        End
        'Dim find As Integer = -5 '23 '110 '14 '-6 '101 '55 '-4 '14 '105 '54
        Dim find As Integer = 89 '30 '54

        Dim findItem As Integer = 0

        startTime()
        For i = 100 To 10000000 '3.4 seconds for 10,000,000 times looking into a 100,000,000 integer array
            'findItem = BiSearch_Int_GTE(a, find)
            findItem = BiSearch_Int_GTE(a, i)
            If i <> findItem Then
                i = i
            End If
            'Debug.Print(findItem)
        Next
        endTime()
        'DiffTime()
        End

        'findItem = BiSearch_Int_GT(a, find)
        'Debug.Print(findItem)

        'findItem = BiSearch_Int_LT(a, find)
        'Debug.Print(findItem)

        'findItem = BiSearch_Int_EQUAL(a, find)
        'Debug.Print(findItem)

        'using Base 0 (array starting with 0)
        'Debug.Print(BiSearch_Int_GTE(a, find, 0) & " " & BiSearch_Int_GT(a, find, 0) & " " & (BiSearch_Int_LT(a, find, 0)) & " " & (BiSearch_Int_EQUAL(a, find, 0)) & " " & (BiSearch_Int_LTE(a, find, 0)))
    End Sub
    Sub setSearchArray_integer(ByRef a() As Integer)
        Dim i As Integer
        Dim nItems As Integer = 100000000 '1000000000 '100
        ReDim a(nItems)
        For i = 1 To nItems
            a(i) = i
        Next

        For i = 30 To 54 : a(i) = 54 : Next
        '>=                              >   50 25 37 31 28 29 30 30 - low = half + 1 Else high = half - 1
        '>                              >=  50 75 62 56 53 54 55 55
        '>=                              <=  50 25 37 31 28 29 30 30 - high = half - 1 Else low = half + 1  note: same as > and low = half + 1 Else high = half - 1)
        '>                              <   50 75 62 56 53 54 55 55 - high = half - 1 Else low = half + 1  note: same as >=    high = half - 1 Else low = half + 1

        '> returns first occurence of FIND or first Item GreaterThan Find
        '< returns                            first Item GreaterThan Find

        '50 75 62 56 53 54 54 
    End Sub
    Function BiSearch_Int(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        Do
            half = (high + low) \ 2
            If Find <= a(half) Then high = half - 1 Else low = half + 1
            Debug.Print(a(half).ToString)
        Loop Until low > high
        Return low
    End Function
    Function BiSearch_Int_GreaterThan(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer

        Dim x As String = ""

        Do
            half = (high + low) \ 2
            'If Find <= a(half) Then high = half - 1 Else low = half + 1 'find lowest = or greaterThan
            'If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest = or greaterThan
            'If Find >= a(half) Then low = half + 1 Else high = half - 1 'find greaterThan

            If Find < a(half) Then high = half - 1 Else low = half + 1 'find greaterThan

            x &= CStr(half) & " "
        Loop Until low > high
        x &= CStr(low) & " "
        Debug.Print(x)
        Return low
    End Function
    Function BiSearch_Int_GTE(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'GTE means GreaterThanOrEqualTo
        'Returns ElementNumber in array - Not the array element itself

        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest EQUAL ITEM is duplicates or greaterThan
            Loop Until low > high
        Else
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM is duplicates or greaterThan
            Loop Until low > high
            If low > base Then
                If Find = a(low - 1) Then low = low - 1
            End If

        End If

        Return low
    End Function
    Function BiSearch_Long_GTE(ByRef a() As Long, ByVal Find As Long, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Long 'GTE means GreaterThanOrEqualTo

        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            'if = exists then this return the lowest =
            'the loop finds a. the lowest = or b. GT
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'find first lowest item that is = to Find OR first Item greaterThan FIND
            Loop Until low > high
        Else
            'if = exists then returns the highest = else >
            'the following loop always returns GT
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
            Loop Until low > high
            If low > base Then
                If Find = a(low - 1) Then low = low - 1
            End If

        End If

        Return low


        '===
        ''Returns ElementNumber in array - Not the array element itself

        'Dim high As Integer = UBound(a)
        'Dim low As Integer = base
        'Dim half As Integer
        'If HightestEqual = False Then
        '    Do
        '        half = (high + low) \ 2
        '        If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest EQUAL ITEM is duplicates or greaterThan
        '    Loop Until low > high
        'Else
        '    Do
        '        half = (high + low) \ 2
        '        If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM is duplicates or greaterThan
        '    Loop Until low > high
        'End If
        'Return low
    End Function
    Function BiSearch_Long_GT(ByRef a() As Long, ByVal Find As Long, Optional ByVal base As Integer = 1) As Integer 'GT means GreaterThan
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        Do
            half = (high + low) \ 2
            If Find < a(half) Then high = half - 1 Else low = half + 1 'find greaterThan
        Loop Until low > high
        Return low
    End Function


    Function BiSearch_Int_GT(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer 'GT means GreaterThan
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        'Dim half2 As Integer
        Do
            half = (high + low) \ 2
            'half2 = low + (high - low) / 2
            'note: if \ 2 is used instead of / 2 then the result is the same as (H+L)\2
            'If half <> half2 Then
            '    Debug.Print(high - low)
            '    Debug.Print(59 / 2)
            '    Debug.Print(59 \ 2)

            '    half = half
            'End If
            If Find < a(half) Then high = half - 1 Else low = half + 1 'find greaterThan
        Loop Until low > high
        Return low
        'You may also wonder as to why mid is calculated using mid = lo + (hi-lo)/2
        'instead of the usual mid = (lo+hi)/2. 
        'This is to avoid another potential rounding bug: in the first case, we want the division
        'to always round down, towards the lower bound. 
        'But division truncates, so when lo+hi would be negative, 
        'it would start rounding towards the higher bound. 
        'Coding the calculation this way ensures that the number divided is always positive
        'and hence always rounds as we want it to. Although the bug doesn’t surface
        'when the search space consists only of positive integers or real numbers, 
        'I've decided to code it this way throughout the article for consistency.
    End Function
    Function BiSearch_Int_LT(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer 'LT means LessThan
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        Do
            half = (high + low) \ 2
            If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest = or greaterThan
        Loop Until low > high
        Return low - 1 'could be 0 or -1 - watch out
    End Function
    Function BiSearch_Int_EQUAL(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'E means EqualTo -Return of -1 means COULDN'T FIND
        Dim HH As Integer = UBound(a)
        Dim high As Integer = HH
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            'if = exists then this return the lowest =
            'the loop finds a. the lowest = or b. GT
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'find first lowest item that is = to Find OR first Item greaterThan FIND
            Loop Until low > high
            If low > HH Then 'must be done this way because LOW must be with array limits
                low = -1
            ElseIf Find = a(low) Then
                'Return low
            Else
                low = -1
            End If
        Else
            'if = exists then this return the highest =
            'the following loop always returns GT
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
            Loop Until low > high
            'now decrement by 1
            low = low - 1
            If Find <> a(low) Then low = -1

        End If

        Return low
    End Function
    Function BiSearch_Int_LTE(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'LTE means LessThanOrEqualTo -Return of -1 means COULDN'T FIND
        Dim HH As Integer = UBound(a)
        Dim high As Integer = HH
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'GTE
            Loop Until low > high
            'above find the lowest = or >
            'If low <= HH Then
            '    If a(low) = Find Then
            '    ElseIf low > base Then
            '        low -= 1
            '    Else
            '        low = -1
            '    End If
            'End If
        Else
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
            Loop Until low > high
        End If
        If low > HH Then 'must be done this way because LOW must be with array limits
            Return low - 1
        ElseIf Find = a(low) Then
            Return low
        ElseIf low > base Then
            Return low - 1
        Else
            Return -1 'error
        End If
    End Function



    'Function BinarySearch(arr As Variant, search As Variant, Optional lastEl As Variant) As Long
    '        ' Binary search in an array of any type
    '        ' Returns the index of the matching item, or -1 if the search fails
    '        '
    '        ' The arrays *must* be sorted, in ascending or descending
    '        ' order (the routines finds out the sort direction).
    '        ' LASTEL is the index of the last item to be searched, and is
    '        ' useful if the array is only partially filled.
    '        '
    '        ' Works with any kind of array, including objects if your are searching 
    '        ' for their default property, and excluding UDTs and fixed-length strings.
    '        ' String are compared in case-sensitive mode.
    '        '
    '        ' You can write faster procedures if you modify the first line
    '        ' to account for a specific data type, eg.
    '        '   Function BinarySearchL (arr() As Long, search As Long,
    '        '  Optional lastEl As Variant) As Long


    '        Dim index As Long
    '        Dim first As Long
    '        Dim last As Long
    '        Dim middle As Long
    '        Dim inverseOrder As Boolean

    '        ' account for optional arguments
    '        If IsMissing(lastEl) Then lastEl = UBound(arr)

    '        first = LBound(arr)
    '        last = lastEl

    '        ' deduct direction of sorting
    '        inverseOrder = (arr(first) > arr(last))

    '        ' assume searches failed
    '        BinarySearch = first - 1

    '        Do
    '            middle = (first + last) \ 2
    '            If arr(middle) = search Then
    '                BinarySearch = middle
    '                Exit Do
    '            ElseIf ((arr(middle) < search) Xor inverseOrder) Then
    '                first = middle + 1
    '            Else
    '                last = middle - 1
    '            End If
    '        Loop Until first > last
    '    End Function
    'Another One
    'Function binary_search(array() As Integer, value As Integer, lo As Integer, hi As Integer) As Integer
    '    Dim middle As Integer

    '    If hi < lo Then
    '        binary_search = 0
    '    Else
    '        middle = (hi + lo) / 2
    '        Select Case value
    '            Case Is < array(middle)
    '                binary_search = binary_search(array(), value, lo, middle - 1)
    '            Case Is > array(middle)
    '                binary_search = binary_search(array(), value, middle + 1, hi)
    '            Case Else
    '                binary_search = middle
    '        End Select
    '    End If
    'End Function

    'Function BiSearch_GTE_FileParts(ByRef a() As fileParts, ByVal Find As String, Optional ByVal base As Integer = 1) As Integer 'GTE means GreaterThanOrEqualTo
    '    'Returns ElementNumber in array - Not the array element itself
    '    Dim xLen As Integer = Len(Find)

    '    Dim high As Integer = UBound(a)
    '    Dim low As Integer = base
    '    Dim half As Integer
    '    Do
    '        half = (high + low) \ 2
    '        If Find > UCase(Mid(a(half).name, 1, xLen)) Then low = half + 1 Else high = half - 1 'find lowest = or greaterThan
    '    Loop Until low > high
    '    Return low
    'End Function
    '===============ret actual element
    Function BiSearch_Int_GTE_RetItem(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'GTE means GreaterThanOrEqualTo
        'Returns ElementNumber in array - Not the array element itself

        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest EQUAL ITEM is duplicates or greaterThan
            Loop Until low > high
        Else
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM is duplicates or greaterThan
            Loop Until low > high
        End If

        Return low
    End Function
    Function BiSearch_Int_GT_RetItem(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer 'GT means GreaterThan
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        Do
            half = (high + low) \ 2


        Loop Until low > high
        Return low
    End Function
    Function BiSearch_Int_LT_RetItem(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1) As Integer 'LT means LessThan
        Dim high As Integer = UBound(a)
        Dim low As Integer = base
        Dim half As Integer
        Do
            half = (high + low) \ 2
            If Find > a(half) Then low = half + 1 Else high = half - 1 'find lowest = or greaterThan
        Loop Until low > high
        If low > base Then
            Return a(low - 1)
        Else
            Return -1
        End If
        '??Return low - 1 'could be 0 or -1 - watch out
    End Function
    Function BiSearch_Int_EQUAL_RetItem(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'E means EqualTo -Return of -1 means COULDN'T FIND
        Dim HH As Integer = UBound(a)
        Dim high As Integer = HH
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'find first lowest item that is = to Find OR first Item greaterThan FIND
            Loop Until low > high
        Else
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
            Loop Until low > high
        End If

        If low > HH Then 'must be done this way because LOW must be with array limits
            Return -1
        ElseIf Find = a(low) Then
            Return a(low)
        Else
            Return -1
        End If

    End Function
    Function BiSearch_Int_LTE_RetItem(ByRef a() As Integer, ByVal Find As Integer, Optional ByVal base As Integer = 1, Optional ByVal HightestEqual As Boolean = False) As Integer 'LTE means LessThanOrEqualTo -Return of -1 means COULDN'T FIND
        Dim HH As Integer = UBound(a)
        Dim high As Integer = HH
        Dim low As Integer = base
        Dim half As Integer
        If HightestEqual = False Then
            Do
                half = (high + low) \ 2
                If Find > a(half) Then low = half + 1 Else high = half - 1 'GTE
            Loop Until low > high
        Else
            Do
                half = (high + low) \ 2
                If Find < a(half) Then high = half - 1 Else low = half + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
            Loop Until low > high
        End If
        If low > HH Then 'must be done this way because LOW must be with array limits
            Return a(low - 1)
        ElseIf Find = a(low) Then
            Return a(low)
        ElseIf low > base Then
            Return a(low - 1)
        Else
            Return -1 'error
        End If
    End Function
    Sub QuickSort_Prototype(ByRef c() As Integer, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Loag
        Dim T, MidValue As Integer 'String or Integer

        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do
            While (c(Low) < MidValue) : Low += 1 : End While
            While (c(High) > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_Prototype(c, First, High)
        If Low < Last Then QuickSort_Prototype(c, Low, Last)
    End Sub
    Sub ShellSort(ByRef list() As Integer, ByVal NumofElements As Integer)
        Dim offset, limit As Integer
        Dim tmp As Integer
        Dim Switch As Boolean

        offset = CInt(NumofElements / 2)
        Do While offset > 0
            limit = NumofElements - offset
            Do
                Switch = False
                For Loop1 = 0 To limit
                    If list(Loop1) > list(Loop1 + offset) Then
                        tmp = list(Loop1)
                        list(Loop1) = list(Loop1 + offset)
                        list(Loop1 + offset) = tmp
                        Switch = True
                        limit = Loop1 - offset
                    End If
                Next Loop1
            Loop While Switch
            offset = CInt(offset / 2)
        Loop

    End Sub
    Sub ShellSortString(ByRef sort() As String, ByVal NumofElements As Integer, Optional ByVal DecendingOrder As Boolean = False)
        'The ShellSort procedure sorts the elements of the array sort() in descending(=1) or ascending(=2) order'and returns this array to the calling procedure

        Dim temp As String
        Dim i, j, span As Integer

        span = NumofElements \ 2
        If DecendingOrder = False Then 'ascending order
            Do While span > 0
                For i = span To NumofElements - 1
                    For j = (i - span + 1) To 1 Step -span
                        If sort(j) >= sort(j + span) Then Exit For
                        temp = sort(j)
                        sort(j) = sort(j + span)
                        sort(j + span) = temp
                    Next j
                Next i
                span = span \ 2
            Loop
        Else 'descending order
            Do While span > 0
                For i = span To NumofElements - 1
                    For j = (i - span + 1) To 1 Step -span
                        'Select Case order
                        '    Case 1
                        '        '1 means descending
                        '        If sort(j) <= sort(j + span) Then Exit For
                        '    Case 2
                        '        '2 means ascending
                        '        If sort(j) >= sort(j + span) Then Exit For
                        'End Select
                        If sort(j) <= sort(j + span) Then Exit For
                        temp = sort(j)
                        sort(j) = sort(j + span)
                        sort(j + span) = temp
                    Next j
                Next i
                span = span \ 2
            Loop
        End If
    End Sub
    'Sub testQsort() 'uses TEST as keyword
    '    'Debug.Print(UBound(gFunctionParts))
    '    'Debug.Print(UBound(gSubParts)) '548  548 000
    '    'End

    '    Dim cA(548000) As String
    '    Dim a(548000) As SubPartsType
    '    'Dim b(548000) As SubPartsType
    '    Dim b(0) As SubPartsType
    '    Dim c As Integer = UBound(gSubParts)
    '    Dim i, j As Integer
    '    For i = 0 To 999
    '        For j = 1 To c
    '            a(j + c * i) = gSubParts(j)
    '            cA(j + c * i) = gSubParts(j).SortBy
    '        Next
    '    Next



    '    'Array.Copy(source, target, target.Length)
    '    'Array.Copy(a, b, a.Length)

    '    b = a.Clone

    '    '''''a(5).SortBy = "A Don"
    '    'Debug.Print(a(5).SortBy & "   " & b(5).SortBy & "   " & b.Length)

    '    StartTime() '.001 seconds


    '    'qSort_SubPartsType(1, a.Count - 1, a) 'a.length

    '    'qSort_SubPartsType(1, a.Length - 1, a)
    '    'QuickSort_SubPartsType(a, 1, a.Length - 1)

    '    QuickSort_String(cA, 1, a.Length - 1)


    '    'testQsort()
    '    'End

    '    EndTime()
    '    DiffTime()

    '    'For i = 1 To a.Length - 1
    '    '    If a(i).SortBy <> b(i).SortBy Then
    '    '        Debug.Print(i & "   " & a(i).SortBy & "   " & b(i).SortBy)
    '    '        'i = i
    '    '    End If
    '    'Next

    '    'For i = 2 To a.Length - 1
    '    '    If a(i).SortBy < a(i - 1).SortBy Then
    '    '        i = i
    '    '    End If
    '    'Next

    '    For i = 2 To a.Length - 1
    '        If cA(i) < cA(i - 1) Then
    '            i = i
    '        End If
    '    Next


    '    End


    '    'Array.Copy(source, target, target.Length)
    '    'MyIntegerA = MyIntegerB.Clone()

    '    'Debug.Print(a(1).SortBy)
    '    'Debug.Print(a(548000).SortBy)
    '    'Debug.Print(a(548).SortBy)
    '    'End


    '    'ReDim Preserve gFunctionParts(FunctionParts_Cnt)


    '    'ReDim Preserve gSubParts(SubParts_Cnt)

    '    'Dim a(UBound(gFunctionParts)) As SubPartsType
    '    'Dim b(UBound(gSubParts)) As SubPartsType
    '    'Dim k As Integer
    '    'For k = 1 To UBound(gFunctionParts)
    '    '    a(k) = gFunctionParts(k)
    '    'Next
    '    'For k = 1 To UBound(gSubParts)
    '    '    b(k) = gSubParts(k)
    '    'Next



    '    'StartTime() '.001 seconds

    '    'qSort_SubPartsType(1, FunctionParts_Cnt, gFunctionParts)
    '    'qSort_SubPartsType(1, SubParts_Cnt, gSubParts)

    '    'testQsort()
    '    'End

    '    '20140626 NEW
    '    'QuickSort_SubPartsType(gFunctionParts, 1, FunctionParts_Cnt)
    '    'QuickSort_SubPartsType(gSubParts, 1, SubParts_Cnt)

    '    'QuickSort_SubPartsType(a, 1, FunctionParts_Cnt)
    '    'QuickSort_SubPartsType(b, 1, SubParts_Cnt)
    '    'For k = 1 To UBound(gFunctionParts)
    '    '    If a(k).SortBy <> gFunctionParts(k).SortBy Then
    '    '        k = k
    '    '    End If
    '    'Next
    '    'For k = 1 To UBound(gSubParts)
    '    '    If b(k).SortBy <> gSubParts(k).SortBy Then
    '    '        k = k
    '    '    End If
    '    'Next



    '    'EndTime()
    '    'DiffTime()


    'End Sub
    'Sub QuickSort_SubPartsType(ByRef c() As SubPartsType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '    If First >= Last Then Exit Sub

    '    Dim T As SubPartsType '==added to eliminate calling swap
    '    Dim Low As Long, High As Long
    '    Dim MidValue As String

    '    Low = First : High = Last

    '    MidValue = c((Low + High) \ 2).SortBy
    '    Do
    '        While (c(Low).SortBy < MidValue) : Low += 1 : End While
    '        While (c(High).SortBy > MidValue) : High -= 1 : End While
    '        If Low <= High Then
    '            T = c(Low) : c(Low) = c(High) : c(High) = T '=='Swap_SubPartsType(c(Low), c(High)) 'replace by left
    '            Low += 1 : High -= 1
    '        End If
    '    Loop While Low <= High
    '    If First < High Then QuickSort_SubPartsType(c, First, High)
    '    If Low < Last Then QuickSort_SubPartsType(c, Low, Last)
    'End Sub
    'Private Sub Swap_SubPartsType(ByRef A As SubPartsType, ByRef B As SubPartsType)
    '    Dim T As SubPartsType = A
    '    A = B
    '    B = T
    'End Sub
    'Sub QuickSort_String_Upper(ByRef a() As String, ByVal First As Integer, ByVal Last As Integer,Optional ByRef CallCount As Integer =0, Optional ByRef c() As StringPointerType ) 'from QuickSort_String
    Sub QuickSort_String_Upper(ByRef a() As String, ByVal First As Integer, ByVal Last As Integer)
        If First >= Last Then Exit Sub

        Dim c(Last) As StringPointerType
        Dim i As Integer
        For i = First To Last : c(i).str = a(i).ToUpper() : c(i).ptr = i : Next
        'QuickSort_StringPointer(c, First, Last) '.4 sec
        Lam(c)
        For i = First To Last : c(i).str = a(c(i).ptr) : Next
        For i = First To Last : a(i) = c(i).str : Next
        ReDim c(0)
    End Sub
    Sub QuickSort_String_Upper2(ByRef a() As String, ByVal First As Integer, ByVal Last As Integer, Optional ByRef CallCount As Integer = 0, Optional ByRef c() As StringPointerType = Nothing) 'from QuickSort_String
        If First >= Last Then Exit Sub

        Dim i As Integer
        If CallCount = 0 Then
            ReDim c(Last)
            For i = First To Last : c(i).str = UCase(a(i)) : c(i).ptr = i : Next
        End If
        CallCount += 1
        '==
        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As String
        Dim T As StringPointerType

        Low = First : High = Last : MidValue = c((Low + High) \ 2).str
        Do
            While (c(Low).str < MidValue) : Low += 1 : End While
            While (c(High).str > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_String_Upper2(a, First, High, CallCount, c)
        If Low < Last Then QuickSort_String_Upper2(a, Low, Last, CallCount, c)
        '===
        CallCount -= 1
        If CallCount = 0 Then
            For i = First To Last : c(i).str = a(c(i).ptr) : Next
            For i = First To Last : a(i) = c(i).str : Next
        End If
    End Sub
    Sub QuickSort_String_Upper3(ByRef a() As String, ByVal First As Integer, ByVal Last As Integer, Optional ByRef CallCount As Integer = 0, Optional ByRef c() As StringPointerType = Nothing) 'from QuickSort_String
        If First >= Last Then Exit Sub

        Dim i As Integer
        If CallCount = 0 Then
            ReDim c(Last)
            For i = First To Last : c(i).str = UCase(a(i)) : c(i).ptr = i : Next
        End If
        CallCount += 1
        '==
        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As String
        Dim T As Integer 'StringPointerType

        Low = First : High = Last : MidValue = c(c((Low + High) \ 2).ptr).str
        Do
            While (c(c(Low).ptr).str < MidValue) : Low += 1 : End While
            While (c(c(High).ptr).str > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low).ptr : c(Low).ptr = c(High).ptr : c(High).ptr = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_String_Upper3(a, First, High, CallCount, c)
        If Low < Last Then QuickSort_String_Upper3(a, Low, Last, CallCount, c)
        '===
        CallCount -= 1
        If CallCount = 0 Then
            For i = First To Last : c(i).str = a(c(i).ptr) : Next
            For i = First To Last : a(i) = c(i).str : Next
        End If
    End Sub
    Sub QuickSort_String(ByRef c() As String, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Loag
        Dim T, MidValue As String 'String or Integer

        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do
            While (c(Low) < MidValue) : Low += 1 : End While
            While (c(High) > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_String(c, First, High)
        If Low < Last Then QuickSort_String(c, Low, Last)
    End Sub
    Sub QuickSort_StringPointer(ByRef c() As StringPointerType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As String 'String or Integer
        Dim T As StringPointerType

        'Debug.Print(First & "  " & Last)

        Low = First : High = Last : MidValue = c((Low + High) \ 2).str
        Do
            While (c(Low).str < MidValue) : Low += 1 : End While
            While (c(High).str > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_StringPointer(c, First, High)
        If Low < Last Then QuickSort_StringPointer(c, Low, Last)
    End Sub
    Sub QuickSort_StringPointer_OnlyPtrChanged(ByRef c() As StringPointerType, ByVal First As Integer, ByVal Last As Integer)
        'case insensitive
        'created 2/28/2015 -dsc
        If First >= Last Then Exit Sub
        'String.Equals(str1, str2, StringComparison.InvariantCultureIgnoreCase
        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As String 'String or Integer
        Dim T As Integer 'StringPointerType

        'Debug.Print(First & "  " & Last)

        Low = First : High = Last : MidValue = c(c((Low + High) \ 2).ptr).str
        Do
            'While (c(r(Low)).str < MidValue) : Low += 1 : End While
            'While (c(r(High)).str > MidValue) : High -= 1 : End While
            While StrComp(c(c(Low).ptr).str, MidValue, CompareMethod.Text) < 0 : Low += 1 : End While
            While StrComp(c(c(High).ptr).str, MidValue, CompareMethod.Text) > 0 : High -= 1 : End While

            If Low <= High Then T = c(Low).ptr : c(Low).ptr = c(High).ptr : c(High).ptr = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_StringPointer_OnlyPtrChanged(c, First, High)
        If Low < Last Then QuickSort_StringPointer_OnlyPtrChanged(c, Low, Last)
    End Sub

    Sub QuickSort_StringPointer_ReturnPointerArray(ByRef c() As StringPointerType, ByVal First As Integer, ByVal Last As Integer, ByRef r() As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        'case insensitive
        If First >= Last Then Exit Sub
        'String.Equals(str1, str2, StringComparison.InvariantCultureIgnoreCase
        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As String 'String or Integer
        Dim T As Integer 'StringPointerType

        'Debug.Print(First & "  " & Last)

        Low = First : High = Last : MidValue = c(r((Low + High) \ 2)).str
        Do
            'While (c(r(Low)).str < MidValue) : Low += 1 : End While
            'While (c(r(High)).str > MidValue) : High -= 1 : End While
            While StrComp(c(r(Low)).str, MidValue, CompareMethod.Text) < 0 : Low += 1 : End While
            While StrComp(c(r(High)).str, MidValue, CompareMethod.Text) > 0 : High -= 1 : End While

            If Low <= High Then T = r(Low) : r(Low) = r(High) : r(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_StringPointer_ReturnPointerArray(c, First, High, r)
        If Low < Last Then QuickSort_StringPointer_ReturnPointerArray(c, Low, Last, r)
    End Sub
    Sub QuickSort_String_ReturnPointerArray(ByRef c() As String, ByVal First As Integer, ByVal Last As Integer, ByRef r() As Integer)
        'case insensitive
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Loag
        Dim MidValue As String 'String or Integer
        Dim T As Integer 'StringPointerType

        Low = First : High = Last
        MidValue = c(r((Low + High) \ 2))
        Do
            'While (c(r(Low)) < MidValue) : Low += 1 : End While
            'While (c(r(High)) > MidValue) : High -= 1 : End While

            While StrComp(c(r(Low)), MidValue, CompareMethod.Text) < 0 : Low += 1 : End While
            While StrComp(c(r(High)), MidValue, CompareMethod.Text) > 0 : High -= 1 : End While


            If Low <= High Then T = r(Low) : r(Low) = r(High) : r(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_String_ReturnPointerArray(c, First, High, r)
        If Low < Last Then QuickSort_String_ReturnPointerArray(c, Low, Last, r)
    End Sub

    Sub QuickSort_String_ignoreCase(ByRef c() As String, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Loag
        Dim T, MidValue As String 'String or Integer

        'Low = First : High = Last : MidValue = UCase(c((Low + High) \ 2))
        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do

            'While UCase(c(Low)) < MidValue : Low += 1 : End While
            'While UCase(c(High)) > MidValue : High -= 1 : End While


            'While String.Compare(c(Low), MidValue, True) < 0 : Low += 1 : End While
            'While String.Compare(c(High), MidValue, True) > 0 : High -= 1 : End While

            While StrComp(c(Low), MidValue, CompareMethod.Text) < 0 : Low += 1 : End While
            While StrComp(c(High), MidValue, CompareMethod.Text) > 0 : High -= 1 : End While

            'While StrComp(c(Low), MidValue, CompareMethod.Binary) < 0 : Low += 1 : End While
            'While StrComp(c(High), MidValue, CompareMethod.Binary) > 0 : High -= 1 : End While


            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_String_ignoreCase(c, First, High)
        If Low < Last Then QuickSort_String_ignoreCase(c, Low, Last)
    End Sub
    ' Sub MergeSort()
    ' Omit pvarMirror, plngLeft & plngRight; they are used internally during recursion
    Sub MergeSort1(ByRef aArray() As String, ByVal Left As Integer, ByVal Right As Integer, Optional ByRef MirrorArray() As String = Nothing)

        Dim MidValue, L, R, O As Integer
        Dim T As String

        If MirrorArray Is Nothing Then
            'If MirrorArray.Length = 0 Then
            ReDim MirrorArray(Right)
            For L = Left To Right
                aArray(L) = UCase(aArray(L))
            Next
        End If

        MidValue = Right - Left
        If MidValue = 1 Then
            If aArray(Left) > aArray(Right) Then T = aArray(Left) : aArray(Left) = aArray(Right) : aArray(Right) = T
        ElseIf MidValue > 1 Then
            MidValue = MidValue \ 2 + Left
            MergeSort1(aArray, Left, MidValue, MirrorArray)
            MergeSort1(aArray, MidValue + 1, Right, MirrorArray)

            ' Merge the resulting halves
            L = Left ' start of first (left) half
            R = MidValue + 1 ' start of second (right) half
            O = Left ' start of output (mirror array)

            Do
                If aArray(R) < aArray(L) Then
                    MirrorArray(O) = aArray(R) : R = R + 1
                    If R > Right Then
                        For L = L To MidValue : O = O + 1 : MirrorArray(O) = aArray(L) : Next
                        Exit Do
                    End If
                Else
                    MirrorArray(O) = aArray(L) : L = L + 1
                    If L > MidValue Then
                        For R = R To Right : O = O + 1 : MirrorArray(O) = aArray(R) : Next
                        Exit Do
                    End If
                End If
                O = O + 1
            Loop
            For O = Left To Right : aArray(O) = MirrorArray(O) : Next
        End If
    End Sub
    Sub IntegerSortOneMil(ByRef a() As Integer, Optional ByVal base As Integer = 1) 'array can't contain a number > 1,000,000 or OneMil '500,000 numbers .02 seconds
        Dim i, j, k, nItems As Integer
        nItems = UBound(a)
        k = base - 1
        Dim b(1000000) As Integer
        For i = base To nItems
            If a(i) > 1000000 Then
                QuickSort_Integer(a, base, nItems)
                Exit Sub
            End If
            b(a(i)) += 1
        Next
        For i = 0 To 1000000 'base To nItems
            For j = 1 To b(i) : k += 1 : a(k) = i : Next
        Next
    End Sub
    Sub QuickSort_Integer(ByRef c() As Integer, ByVal First As Integer, ByVal Last As Integer)
        If First >= Last Then Exit Sub

        Dim Low, High As Integer
        Dim T, MidValue As Integer

        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do
            While (c(Low) < MidValue) : Low += 1 : End While
            While (c(High) > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_Integer(c, First, High)
        If Low < Last Then QuickSort_Integer(c, Low, Last)
    End Sub
    Sub QuickSort_Decimal(ByRef c() As Decimal, ByVal First As Integer, ByVal Last As Integer)
        If First >= Last Then Exit Sub

        Dim Low, High As Integer
        Dim T, MidValue As Decimal 'Integer

        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do
            While (c(Low) < MidValue) : Low += 1 : End While
            While (c(High) > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_Decimal(c, First, High)
        If Low < Last Then QuickSort_Decimal(c, Low, Last)
    End Sub
    Sub QuickSort_Long(ByRef c() As Long, ByVal First As Integer, ByVal Last As Integer)
        If First >= Last Then Exit Sub

        Dim Low, High As Integer
        Dim T, MidValue As Long

        Low = First : High = Last : MidValue = c((Low + High) \ 2)
        Do
            While (c(Low) < MidValue) : Low += 1 : End While
            While (c(High) > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_Long(c, First, High)
        If Low < Last Then QuickSort_Long(c, Low, Last)
    End Sub
    Sub QuickSort_Decimal_Hybred(ByRef c() As Decimal, ByVal First As Integer, ByVal Last As Integer) 'twice as fast as 'QuickSort_Decimal' but twice the memory 2 arrays!
        If First >= Last Then Exit Sub

        Dim b(Last) As Long
        Dim i As Integer
        For i = First To Last : b(i) = CLng(c(i) * 100) : Next
        QuickSort_Long(b, First, Last)
        For i = First To Last : c(i) = CDec(b(i) / 100) : Next
        ReDim b(0)
    End Sub
    Sub Lam(ByRef c() As StringPointerType)
        'Array.Sort(c, (Function(x As StringPointerType, y As StringPointerType) String.Compare(x.str, y.str)))
        startTime()
        System.Array.Sort(c, (Function(x As StringPointerType, y As StringPointerType) String.Compare(x.str, y.str)))
        'Array.Sort(c)
        endTime()
        'DiffTime()
    End Sub
    Sub QuickSort_TwoInteger_i2(ByRef c() As TwoIntegersType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
        If First >= Last Then Exit Sub

        Dim Low, High As Integer 'Always Integer or Long
        Dim MidValue As Integer
        Dim T As TwoIntegersType

        Low = First : High = Last : MidValue = c((Low + High) \ 2).i2
        Do
            While (c(Low).i2 < MidValue) : Low += 1 : End While
            While (c(High).i2 > MidValue) : High -= 1 : End While
            If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
        Loop While Low <= High
        If First < High Then QuickSort_TwoInteger_i2(c, First, High)
        If Low < Last Then QuickSort_TwoInteger_i2(c, Low, Last)
    End Sub
    Sub test_bisearch2()

        Dim nItems As Integer = 1000000
        Dim a(nItems) As Integer
        Dim i As Integer
        For i = 1 To nItems : a(i) = i
        Next
        For i = 2 To 200
            a(i) = 9
        Next

        Dim s As Integer = 9 '100 '-199 ' 9 '607891
        Dim f As Integer

        'both of the following now work!
        'f = BiSearch_Int_EQUAL(a, s, , True) 'if = returns Highest - ; else -1
        'f = BiSearch_Int_EQUAL(a, s) 'if = returns lowest = ; else -1
        '-----------------
        'f = BiSearch_Int_GTE(a, s, , True) 'returns higest = or GT
        'f = BiSearch_Int_GTE(a, s) 'returns lowest = or GT

        f = BiSearch_Int_GT(a, s)
        'f = BiSearch_Int_LT(a, s, 0)
        If s <> f Then
            Debug.Print(f.ToString)
            End
        End If
        Debug.Print("ok")

        End
    End Sub
    'Sub test_QuickSort_LongInteger_Long()
    '    Dim i As Integer
    '    'ApptGV.ccPath = "\\WDMYCLOUD\SmartWare\TodayPath20151225\"
    '    ApptGV.ccPath = "c:\TodayPath20151225\"
    '    'startTime()
    '    ApptGV.AllAppts_LongInteger = getLongInteger_Apptfil() 'as of 11/11/2015 'take .003 seconds 878 records
    '    Dim a() As LongIntegerType = ApptGV.AllAppts_LongInteger.Clone
    '    'endTime()
    '    'Debug.Print(UBound(ApptGV.AllAppts_LongInteger))

    '    startTime()
    '    QuickSort_LongInteger_Long(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger)) '.001 seconds
    '    endTime()
    '    'If AccendingOrderOK() = False Then
    '    '    i = i
    '    'End If

    '    startTime()
    '    'QuickSort(ApptGV.AllAppts_LongInteger, 1, UBound(ApptGV.AllAppts_LongInteger))
    '    QuickSort_LongInteger_Long_Descending(ApptGV.AllAppts_LongInteger, 0, UBound(ApptGV.AllAppts_LongInteger)) '.001 seconds
    '    'Dim q = From c In ApptGV.AllAppts_LongInteger Order By c.L Descending Select c
    '    Dim q = From c In a Order By c.L Descending Select c
    '    For i = 0 To ApptGV.AllAppts_LongInteger.Count - 1
    '        If ApptGV.AllAppts_LongInteger(i).L <> q(i).L Then
    '            i = i
    '        End If
    '        If ApptGV.AllAppts_LongInteger(i).i <> q(i).i Then
    '            i = i
    '        End If

    '    Next
    '    'Array.Sort(a, Function(x, y) x.CompareTo(y) * -1)
    '    'Array.Sort(ApptGV.AllAppts_LongInteger, Function(x, y) x.CompareTo(y) * -1)
    '    endTime()
    '    'For i = UBound(ApptGV.AllAppts_LongInteger) To 0 Step -1
    '    '    Debug.Print(i & " " & ApptGV.AllAppts_LongInteger(i).L)
    '    'Next
    '    If DescendingOrderOK() = False Then
    '        For i = 0 To 5
    '            Debug.Print(i & " " & ApptGV.AllAppts_LongInteger(i).L)
    '        Next
    '        i = i
    '    End If
    '    i = UBound(ApptGV.AllAppts_LongInteger)
    '    Debug.Print(ApptGV.AllAppts_LongInteger(i).L)
    '    For i = 5 To 0 Step -1
    '        Debug.Print(i & " " & ApptGV.AllAppts_LongInteger(i).L)
    '    Next


    '    '5 636156000000000000   5 636155784000000000
    '    '4 636166128000000000   4 636156000000000000
    '    '3 636166155000000000   3 636166128000000000
    '    '2 636225924000000000   2 636166155000000000
    '    '1 636293367000000000   1 636225924000000000
    '    '0 0                    0 636293367000000000

    '    'test_QuickSort_LongInteger_Long()
    '    '0
    '    '5 636155784000000000
    '    '4 636156000000000000
    '    '3 636166128000000000
    '    '2 636166155000000000
    '    '1 636225924000000000
    '    '0 636293367000000000


    'End Sub
    'Function AccendingOrderOK() As Boolean
    '    Dim b As Boolean = True
    '    Dim nRecs As Integer = UBound(ApptGV.AllAppts_LongInteger)
    '    Dim i As Integer
    '    For i = 1 To nRecs
    '        If ApptGV.AllAppts_LongInteger(i).L < ApptGV.AllAppts_LongInteger(i - 1).L Then
    '            b = False
    '            Exit For
    '        End If
    '    Next
    '    Return b
    'End Function
    'Function DescendingOrderOK() As Boolean
    '    Dim b As Boolean = True
    '    Dim nRecs As Integer = UBound(ApptGV.AllAppts_LongInteger)
    '    Dim i As Integer
    '    For i = 2 To nRecs
    '        If ApptGV.AllAppts_LongInteger(i).L > ApptGV.AllAppts_LongInteger(i - 1).L Then
    '            b = False
    '            Exit For
    '        End If
    '    Next
    '    Return b
    'End Function
    'Sub QuickSort_ApptRecordsType_Descending(ByRef c() As ApptRecordsType, ByVal First As Integer, ByVal Last As Integer)
    '    'see: QuickSort_LongInteger_Long_Descending for referenct
    '    If First >= Last Then Exit Sub

    '    Dim Low, High As Integer 'Long 'Always Integer or Long
    '    Dim MidValue As Long 'Integer
    '    Dim T As ApptRecordsType 'LongIntegerType

    '    Low = First : High = Last : MidValue = c((Low + High) \ 2).ApptRec.dTics
    '    Do
    '        While (c(Low).ApptRec.dTics > MidValue) : Low += 1 : End While
    '        While (c(High).ApptRec.dTics < MidValue) : High -= 1 : End While

    '        If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '    Loop While Low <= High
    '    If First < High Then QuickSort_ApptRecordsType_Descending(c, First, High)
    '    If Low < Last Then QuickSort_ApptRecordsType_Descending(c, Low, Last)
    'End Sub
    'Sub QuickSort_LongInteger_Long_Descending(ByRef c() As LongIntegerType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '    If First >= Last Then Exit Sub

    '    Dim Low, High As Integer 'Long 'Always Integer or Long
    '    Dim MidValue As Long 'Integer
    '    Dim T As LongIntegerType

    '    Low = First : High = Last : MidValue = c((Low + High) \ 2).L
    '    Do
    '        'For Descending Order Simply change the < and > signs on next Two Lines
    '        'While (c(Low).L < MidValue) : Low += 1 : End While
    '        'While (c(High).L > MidValue) : High -= 1 : End While
    '        While (c(Low).L > MidValue) : Low += 1 : End While
    '        While (c(High).L < MidValue) : High -= 1 : End While

    '        If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '    Loop While Low <= High
    '    If First < High Then QuickSort_LongInteger_Long_Descending(c, First, High)
    '    If Low < Last Then QuickSort_LongInteger_Long_Descending(c, Low, Last)

    'End Sub
    'Sub QuickSort_LongInteger_Long(ByRef c() As LongIntegerType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '    'new 11/26/2016
    '    If First >= Last Then Exit Sub

    '    Dim Low, High As Integer 'Long 'Always Integer or Long
    '    Dim MidValue As Long 'Integer
    '    Dim T As LongIntegerType

    '    Low = First : High = Last : MidValue = c((Low + High) \ 2).L
    '    Do
    '        While (c(Low).L < MidValue) : Low += 1 : End While
    '        While (c(High).L > MidValue) : High -= 1 : End While
    '        If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '    Loop While Low <= High
    '    If First < High Then QuickSort_LongInteger_Long(c, First, High)
    '    If Low < Last Then QuickSort_LongInteger_Long(c, Low, Last)
    'End Sub
    'Sub QuickSort_LongInteger_Int(ByRef c() As LongIntegerType, ByVal First As Integer, ByVal Last As Integer) 'works very well! - eliminates calling the swap routine dsc - 20140627
    '    'new 11/26/2016
    '    If First >= Last Then Exit Sub

    '    Dim Low, High As Integer 'Long 'Always Integer or Long
    '    Dim MidValue As Long 'Integer
    '    Dim T As LongIntegerType

    '    Low = First : High = Last : MidValue = c((Low + High) \ 2).i
    '    Do
    '        While (c(Low).i < MidValue) : Low += 1 : End While
    '        While (c(High).i > MidValue) : High -= 1 : End While
    '        If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '    Loop While Low <= High
    '    If First < High Then QuickSort_LongInteger_Int(c, First, High)
    '    If Low < Last Then QuickSort_LongInteger_Int(c, Low, Last)
    'End Sub
    'Sub QuickSort_ApptRecType_dTics(ByRef c() As ApptRecType, ByVal First As Integer, ByVal Last As Integer)
    '    'new 11/26/2016
    '    If First >= Last Then Exit Sub

    '    Dim Low, High As Integer 'Long 'Always Integer or Long
    '    Dim MidValue As Long 'Integer
    '    Dim T As ApptRecType '//*

    '    Low = First : High = Last : MidValue = c((Low + High) \ 2).dTics  '//*
    '    Do
    '        While (c(Low).dTics < MidValue) : Low += 1 : End While '//*
    '        While (c(High).dTics > MidValue) : High -= 1 : End While '//*
    '        If Low <= High Then T = c(Low) : c(Low) = c(High) : c(High) = T : Low += 1 : High -= 1
    '    Loop While Low <= High
    '    If First < High Then QuickSort_ApptRecType_dTics(c, First, High) '//*
    '    If Low < Last Then QuickSort_ApptRecType_dTics(c, Low, Last) '//*
    'End Sub
#Region "test_search3"

    Sub test_searchDates()
        startTime()
        Dim a() As Date = get50000Dates() '0.002 secs


        Dim search As Date
        Dim i As Integer
        'Dim retVal As Integer
        For i = 1 To UBound(a) ' 50000
            search = a(i)
            'retVal = matchInteger(search, a)

            'E (GE)
            'retVal = bisearch_E(search, a)

            'retVal = search_dates_LTE(search, a)
            'retVal = Search_dates_LT(search, a)

            Dim d As Date
            Dim base As Integer = 1 'vs 0

            d = Search_Dates_LT(search, a, base)
            If d = Nothing Then
                Debug.Print(search.ToString)
            Else
                If d <> DateAdd(DateInterval.Day, -1, search) Then 'E (actually GE (greaterThan or Equal to))
                    'If retVal <> (search + 1) Then 'GT (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
                    i = i
                End If

            End If
            '===============
            d = search_Dates_GT(search, a)
            If d = Nothing Then
                Debug.Print(search.ToString)
            Else
                If d <> DateAdd(DateInterval.Day, 1, search) Then 'E (actually GE (greaterThan or Equal to))
                    'If retVal <> (search + 1) Then 'GT (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
                    i = i
                End If

            End If

            '===============

            'If i = 1 Then Debug.Print(i & " " & a(retVal)) '1 3/18/1940   to 50000 2/6/2077 - approx 137 yrs
            'GT
            'retVal = bisearch_GT(search, a)


            'If a(retVal) <> search Then 'E (actually GE (greaterThan or Equal to))
            'If a(retVal) <> DateAdd(DateInterval.Day, -1, search) Then 'E (actually GE (greaterThan or Equal to))
            '    'If retVal <> (search + 1) Then 'GT (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
            '    i = i
            'End If
        Next
        endTime() '0.01 secs for search

    End Sub
    Function bisearch_Dates_E(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Integer '0.0000002 secs
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search > a(mid) Then low = mid + 1 Else high = mid - 1
        Loop Until low > high
        Return low
    End Function


    Sub test_search3()
        test_searchDates()
        End

        'Note: bisearch_E is 190 faster than matchInteger

        'test_search3LessThan()
        'End
        '=========

        'startTime() '0.0110292
        Dim a() As Integer = get50000()
        startTime()
        Dim search As Integer
        Dim i As Integer
        Dim retVal As Integer


        'retVal = bisearch_E(312, a)
        ''25000 1 24999 24998
        ''12500 1 12499 12498
        ''6250 1 6249 6248
        ''3125 1 3124 3123
        ''1562 1 1561 1560
        ''781 1 780 779
        ''390 1 389 388
        ''195 196 389 193
        ''292 293 389 96
        ''341 293 340 47
        ''316 293 315 22
        ''304 305 315 10
        ''310 311 315 4
        ''313 311 312 1
        ''311 312 312 0
        ''312 312 311 -1
        'End

        For i = 1 To 50000
            search = i
            'retVal = matchInteger(search, a)

            'E (GE)
            retVal = bisearch_E(search, a)
            'GT
            'retVal = bisearch_GT(search, a)


            If a(retVal) <> search Then 'E (actually GE (greaterThan or Equal to))
                'If retVal <> (search + 1) Then 'GT (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
                i = i
            End If
        Next
        endTime()
    End Sub
    Function get50000() As Integer()
        Dim i As Integer
        Dim a(50000) As Integer
        For i = 1 To 50000
            a(i) = i
        Next
        Return a

    End Function
    Function get50000Dates() As Date()
        Dim nItems As Integer = 50000 'even at 1,000,000 dates - search is < .0000004 or less than 1 millioneth of a second!
        Dim i As Integer
        Dim a(nItems) As Date
        'Dim startDate As Date = New Date(1940, 3, 17)
        a(0) = New Date(1940, 3, 17)
        For i = 1 To nItems
            a(i) = DateAdd(DateInterval.Day, 1, a(i - 1))
        Next
        Return a
    End Function

    Function matchInteger(ByVal search As Integer, ByRef a() As Integer) As Integer
        Dim hhigh As Integer = UBound(a)
        Dim i As Integer
        Dim hit As Integer = 0
        For i = 1 To hhigh
            If a(i) = search Then
                hit = i
                Exit For
            End If
        Next
        Return hit
    End Function
    Function bisearch_E(ByVal search As Integer, ByRef a() As Integer, Optional ByVal base As Integer = 1) As Integer
        'Static cnt As Decimal = 0
        'Static cnt2 As Decimal = 0
        'cnt2 += 1

        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search > mid Then low = mid + 1 Else high = mid - 1
            'Debug.Print(search & " " & mid & " " & low & " " & high & " " & (high - low))
            'cnt += 1
        Loop Until low > high
        'Debug.Print(cnt / cnt2)'15.6893
        Return low
    End Function
    '25000 1 24999 24998
    '12500 1 12499 12498
    '6250 1 6249 6248
    '3125 1 3124 3123
    '1562 1 1561 1560
    '781 1 780 779
    '390 1 389 388
    '195 196 389 193
    '292 293 389 96
    '341 293 340 47
    '316 293 315 22
    '304 305 315 10
    '310 311 315 4
    '313 311 312 1
    '311 312 312 0
    '312 312 311 -1

    Function bisearch_GT(ByVal search As Integer, ByRef a() As Integer, Optional ByVal base As Integer = 1) As Integer
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search >= mid Then low = mid + 1 Else high = mid - 1
        Loop Until low > high
        Return low
    End Function
    '====
    Sub test_search3LessThan()
        startTime() '0.0110292
        Dim a() As Integer = get50000()
        Dim search As Integer
        Dim i As Integer
        Dim retVal As Integer
        For i = 1 To 50000
            search = i

            'retVal = bisearch_LTE(search, a)
            retVal = BiSearch_Int_LT(a, search)


            'If retVal <> (search) Then 'LTE (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
            If retVal <> (search - 1) Then 'LTE (use for GT instead of: 'If a(retVal) <> (search + 1) Then 'GT)
                i = i
            End If
        Next
        endTime()
    End Sub
    Function search_dates_LTE(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Integer
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        'Dim base As Integer = 1
        Do
            mid = (high + low) \ 2
            If search < a(mid) Then high = mid - 1 Else low = mid + 1 'find greatest EQUAL ITEM if there are duplicate EQUAL ITEMS or greaterThan
        Loop Until low > high
        If low > hhigh Then 'must be done this way because LOW must be with array limits
            Return low - 1
        ElseIf search = a(low) Then
            Return low
        ElseIf low > base Then
            Return low - 1
        Else
            Return -1 'error
        End If

    End Function
    Function Search_Dates_LT(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Date 'works with array of dates that are unique (might work with duplicate dates also but not yet tested for that)
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        'Dim base As Integer = 1
        Do
            mid = (high + low) \ 2
            If search > a(mid) Then low = mid + 1 Else high = mid - 1 'find lowest = or greaterThan
        Loop Until low > high

        If (low - 1) < base Then
            Return Nothing '""
        Else
            Return a(low - 1)
        End If
        'Return low - 1 'could be 0 or -1 - watch out
    End Function

    Function search_Dates_GT(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Date
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search >= a(mid) Then low = mid + 1 Else high = mid - 1
        Loop Until low > high
        Dim retVal As Integer = low
        If low > hhigh Then
            Return Nothing '""
        Else
            Return a(low)
        End If

        'Return low
    End Function
    Function search_Dates_GTi(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Integer 'make sure other routines include BASE optional parameter!
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search >= a(mid) Then low = mid + 1 Else high = mid - 1
        Loop Until low > high
        Return low
    End Function
    Function Search_Dates_LTi(ByVal search As Date, ByRef a() As Date, Optional ByVal base As Integer = 1) As Integer 'works with array of dates that are unique (might work with duplicate dates also but not yet tested for that)
        Dim hhigh As Integer = UBound(a)
        Dim high As Integer = hhigh
        Dim low As Integer = base
        Dim mid As Integer = 0
        Do
            mid = (high + low) \ 2
            If search > a(mid) Then low = mid + 1 Else high = mid - 1 'find lowest = or greaterThan
        Loop Until low > high
        Return low - 1 'could be 0 or -1 - watch out
    End Function


#End Region 'test_search3
End Module
