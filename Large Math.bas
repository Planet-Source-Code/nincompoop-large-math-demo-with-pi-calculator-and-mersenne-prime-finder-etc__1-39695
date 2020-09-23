Attribute VB_Name = "modLargeMath"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Private BitMask(31) As Long
Private HexBits(16) As Integer
Private DidInit As Boolean

Public Sub ZeroArray(Num() As Integer, UBNum As Long)

    ReDim Num(0)
    UBNum = 0

End Sub

Public Sub UnityArray(Num() As Integer, UBNum As Long)

    ReDim Num(0)
    UBNum = 0
    Num(0) = 1

End Sub

Public Function IsZero(Num() As Integer, UBNum As Long) As Boolean

    TrimArray Num, UBNum
    If UBNum = 0 And Num(0) = 0 Then IsZero = True

End Function

Public Function IsUnity(Num() As Integer, UBNum As Long) As Boolean

    TrimArray Num, UBNum
    If UBNum = 0 And Num(0) = 1 Then IsUnity = True

End Function

Public Function IsArrayPrime(Num() As Integer, UBNum As Long) As Boolean

    Dim Div() As Integer
    Dim UBDiv As Long

    IsArrayPrime = IsArrayPrime2(Num, UBDiv, Div, UBDiv)

End Function

Public Function IsArrayPrime2(Num() As Integer, UBNum As Long, Div() As Integer, UBDiv As Long) As Boolean

    Dim Root() As Integer
    Dim UBRoot As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Two(0) As Integer
    Dim FileNo As Integer
    Dim curNum As String

    FileNo = FreeFile
    Open "PrimesR.txt" For Input Access Read Lock Write As FileNo
    Sqrt Num, UBNum, Root, UBRoot
    IsArrayPrime2 = True
    Do While Not EOF(FileNo)
        Input #FileNo, curNum
        StringToArray curNum, Div, UBDiv
        If ArrayCmp(Div, UBDiv, Root, UBRoot) = 1 Then Exit Do
        If ArrayDivide(Num, UBNum, Div, UBDiv, Q, UBQ, R, UBR) Then
            IsArrayPrime2 = False
            Exit Do
        End If
    Loop
    Close #FileNo
    If ArrayCmp(Div, UBDiv, Root, UBRoot) <> 1 And IsArrayPrime2 Then
        Two(0) = 2
        Do While 1
            AddArray2 Div, UBDiv, Two, 0
            If ArrayCmp(Div, UBDiv, Root, UBRoot) = 1 Then Exit Do
            If ArrayDivide(Num, UBNum, Div, UBDiv, Q, UBQ, R, UBR) Then
                IsArrayPrime2 = False
                Exit Do
            End If
        Loop
    End If
    If IsArrayPrime2 Then UnityArray Div, UBDiv

End Function

Public Function IsArrayNinProbablePrime(Num() As Integer, UBNum As Long) As Boolean

    Dim RPower() As Integer
    Dim UBRPower As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Residue() As Integer
    Dim UBResidue As Long
    Dim Bin() As Integer
    Dim UBBin As Long
    Dim tmpArr() As Integer
    Dim UBtmpArr As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim One(0) As Integer
    Dim i As Long

    ArrayToBin Num, UBNum, Bin, UBBin
    i = LenBin(Bin, UBBin)

    StringToArray Format$(i), tmpArr, UBtmpArr
    One(0) = 2
    ArrayDivide Num, UBNum, One, 0, tmpAns, UBtmpAns, R, UBR
    ArrayDivide tmpAns, UBtmpAns, tmpArr, UBtmpArr, Q, UBQ, RPower, UBRPower
    ArrayToBin Q, UBQ, Bin, UBBin
    PowerOf2 i, tmpAns, UBtmpAns
    ReDim Residue(0)
    UBResidue = 0
    Residue(0) = 1
    For i = 1 To LenBin(Bin, UBBin)
        ArrayDivide tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, R, UBR
        If IsNthBitSet(i, Bin, UBBin) Then
            ArrayMultiply Residue, UBResidue, R, UBR, tmpArr, UBtmpArr
            ArrayDivide tmpArr, UBtmpArr, Num, UBNum, Q, UBQ, Residue, UBResidue
        End If
        ArrayMultiply R, UBR, R, UBR, tmpAns, UBtmpAns
    Next i
    ArrayPower One, 0, RPower, UBRPower, R, UBR
    ArrayMultiply Residue, UBResidue, R, UBR, tmpAns, UBtmpAns
    One(0) = 8
    ArrayDivide Num, UBNum, One, 0, Q, UBQ, R, UBR
    One(0) = 1
    If R(0) = 3 Or R(0) = 5 Then
        AddArray2 tmpAns, UBtmpAns, One, 0
    Else
        AddArray2 tmpAns, UBtmpAns, Num, UBNum
        SubArray2 tmpAns, UBtmpAns, One, 0
    End If
    If ArrayDivide(tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, Residue, UBResidue) Then
        IsArrayNinProbablePrime = True
    End If

End Function

Public Function IsArrayStrongProbablePrime(Num() As Integer, UBNum As Long, Base() As Integer, UBBase As Long) As Boolean

    Dim RPower() As Integer
    Dim UBRPower As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Residue() As Integer
    Dim UBResidue As Long
    Dim Bin() As Integer
    Dim UBBin As Long
    Dim tmpArr() As Integer
    Dim UBtmpArr As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim One(0) As Integer
    Dim i As Long
    Dim s As Long

    One(0) = 1
    SubArray Num, UBNum, One, 0, tmpAns, UBtmpAns
    One(0) = 2
    Do While tmpAns(0) Mod 2 = 0
        ArrayDivide tmpAns, UBtmpAns, One, 0, Q, UBQ, R, UBR
        CopyArray Q, UBQ, tmpAns, UBtmpAns
        s = s + 1
    Loop
    If s = 0 Then
        Exit Function
    End If

    CopyArray Base, UBBase, tmpArr, UBtmpArr
    i = 1
    Do While ArrayCmp(tmpArr, UBtmpArr, Num, UBNum) = -1
        i = i + 1
        ArrayMultiply tmpArr, UBtmpArr, Base, UBBase, Q, UBQ
        CopyArray Q, UBQ, tmpArr, UBtmpArr
    Loop
    StringToArray Format$(i), R, UBR
    If ArrayCmp(R, UBR, tmpAns, UBtmpAns) = 1 Then
        CopyArray tmpAns, UBtmpAns, R, UBR
        ArrayPower Base, UBBase, R, UBR, tmpArr, UBtmpArr
    End If
    ArrayDivide tmpAns, UBtmpAns, R, UBR, Q, UBQ, RPower, UBRPower
    ArrayToBin Q, UBQ, Bin, UBBin

    ReDim Residue(0)
    UBResidue = 0
    Residue(0) = 1
    For i = 1 To LenBin(Bin, UBBin)
        ArrayDivide tmpArr, UBtmpArr, Num, UBNum, Q, UBQ, R, UBR
        If IsNthBitSet(i, Bin, UBBin) Then
            ArrayMultiply Residue, UBResidue, R, UBR, tmpAns, UBtmpAns
            ArrayDivide tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, Residue, UBResidue
        End If
        ArrayMultiply R, UBR, R, UBR, tmpArr, UBtmpArr
    Next i
    ArrayPower Base, UBBase, RPower, UBRPower, R, UBR
    ArrayMultiply Residue, UBResidue, R, UBR, tmpAns, UBtmpAns  'tmpAns = a^d

    One(0) = 1
    ArrayDivide tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, R, UBR
    If UBR = 0 And R(0) = 1 Then
        IsArrayStrongProbablePrime = True
        Exit Function
    End If

    For i = 0 To s - 1
        AddArray tmpAns, UBtmpAns, One, 0, tmpArr, UBtmpArr
        If ArrayDivide(tmpArr, UBtmpArr, Num, UBNum, Q, UBQ, R, UBR) Then
            IsArrayStrongProbablePrime = True
            Exit Function
        End If
        ArrayDivide tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, R, UBR
        ArrayMultiply R, UBR, R, UBR, tmpAns, UBtmpAns
    Next i

End Function

Public Function IsMersennePrimeExp(Exp As Long) As Boolean

    Dim Num() As Integer
    Dim UBNum As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim s() As Integer
    Dim UBS As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim One(0) As Integer
    Dim i As Long

    StringToArray Format$(Exp), Num, UBNum
    If Not IsArrayPrime(Num, UBNum) Then
        Exit Function
    End If

    One(0) = 1
    PowerOf2 Exp, Num, UBNum
    SubArray2 Num, UBNum, One, 0

    One(0) = 2
    ReDim s(0)
    s(0) = 4
    UBS = 0
    For i = 3 To Exp
        ArrayMultiply s, UBS, s, UBS, tmpAns, UBtmpAns
        SubArray2 tmpAns, UBtmpAns, One, 0
        ArrayDivide tmpAns, UBtmpAns, Num, UBNum, Q, UBQ, s, UBS
    Next i
    IsMersennePrimeExp = IsZero(s, UBS)

End Function

Public Sub TrimArray(Num() As Integer, UBoundNum As Long)

    Dim nElements As Long

    For nElements = UBoundNum To 0 Step -1
        If Num(nElements) <> 0 Then Exit For
    Next nElements
    If nElements = UBoundNum Then Exit Sub
    If nElements < 0 Then nElements = 0
    ReDim Preserve Num(nElements)
    UBoundNum = nElements

End Sub

Public Sub CopyArray(NumFrom() As Integer, ByVal UBoundNumFrom As Long, NumTo() As Integer, UBoundNumTo As Long)

    ReDim NumTo(UBoundNumFrom)
    UBoundNumTo = UBoundNumFrom
    Do While UBoundNumFrom >= 0
        NumTo(UBoundNumFrom) = NumFrom(UBoundNumFrom)
        UBoundNumFrom = UBoundNumFrom - 1
    Loop

End Sub

Public Sub BinToArray(Bin() As Integer, ByVal UBBin As Long, Num() As Integer, UBNum As Long)

    Dim i As Long
    Dim tmpPow() As Integer
    Dim UBtmpPow As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim Sixteen(0) As Integer
    Dim HDec(0) As Integer

    ReDim Num(0)
    UBNum = 0
    Sixteen(0) = 16
    ReDim tmpPow(0)
    tmpPow(0) = 1
    For i = 0 To UBBin
        If Bin(i) <> 0 Then
            HDec(0) = HexBitsToDec(Bin(i))
            ArrayMultiply tmpPow, UBtmpPow, HDec, 0, tmpAns, UBtmpAns
            AddArray2 Num, UBNum, tmpAns, UBtmpAns
        End If
        ArrayMultiply tmpPow, UBtmpPow, Sixteen, 0, tmpAns, UBtmpAns
        CopyArray tmpAns, UBtmpAns, tmpPow, UBtmpPow
    Next i

End Sub

Public Sub StringToArray(NumStr As String, Num() As Integer, UBoundNum As Long)

    Dim nElements As Long
    Dim Length As Long

    Length = Len(NumStr)
    nElements = (Length - 1) \ 4
    ReDim Num(nElements)
    UBoundNum = nElements
    If Length = 0 Then Exit Sub
    If Length And 3 Then
        Num(nElements) = CInt(Left$(NumStr, Length And 3))
        nElements = nElements - 1
    End If
    Length = Length - 3 - (nElements * 4)
    Do While nElements >= 0
        Num(nElements) = CInt(Mid$(NumStr, Length, 4))
        Length = Length + 4
        nElements = nElements - 1
    Loop

End Sub

Public Function ArrayToString(Num() As Integer, ByVal UBoundNum As Long) As String

    Dim tmpStr As String

    ArrayToString = Format$(Num(UBoundNum), "###0")
    UBoundNum = UBoundNum - 1
    Do While UBoundNum >= 0
        tmpStr = tmpStr & Format$(Num(UBoundNum), "0000")
        If UBoundNum Mod 200 = 0 Then
            ArrayToString = ArrayToString & tmpStr
            tmpStr = ""
        End If
        UBoundNum = UBoundNum - 1
    Loop

End Function

Public Sub ArrayToBin(Num() As Integer, UBNum As Long, Bin() As Integer, UBBin As Long)

    Dim NumArr() As Integer
    Dim UBNumArr As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long
    Dim Sixteen(0) As Integer
    Dim i As Long

    If Not DidInit Then
        Init
    End If
    If IsZero(Num, UBNum) Then
        ReDim Bin(0)
        Bin(0) = 0
        Exit Sub
    End If
    UBBin = (Fix(UBNum * 13.288) + 13) \ 4
    ReDim Bin(UBBin)
    Sixteen(0) = 16
    CopyArray Num, UBNum, NumArr, UBNumArr
    Do While Not (UBNumArr = 0 And NumArr(0) = 0)
        ArrayDivide NumArr, UBNumArr, Sixteen, 0, Q, UBQ, R, UBR
        Bin(i) = HexBits(R(0))
        CopyArray Q, UBQ, NumArr, UBNumArr
        i = i + 1
    Loop
    TrimArray Bin, UBBin

End Sub

Public Function ArrayToHexStr(Num() As Integer, ByVal UBNum As Long) As String

    Dim Bin() As Integer
    Dim UBBin As Long

    ArrayToBin Num, UBNum, Bin, UBBin
    Do
        ArrayToHexStr = ArrayToHexStr & Bin(UBBin)
        UBBin = UBBin - 1
    Loop While UBBin >= 0

End Function

Public Function ArrayCmp(Num1() As Integer, ByVal UBoundNum1 As Long, Num2() As Integer, ByVal UBoundNum2 As Long) As Integer

    If UBoundNum1 > UBoundNum2 Then
        ArrayCmp = 1
        Exit Function
    ElseIf UBoundNum1 < UBoundNum2 Then
        ArrayCmp = -1
        Exit Function
    Else
        Do While UBoundNum1 >= 0
            If Num1(UBoundNum1) > Num2(UBoundNum1) Then
                ArrayCmp = 1
                Exit Function
            ElseIf Num1(UBoundNum1) < Num2(UBoundNum1) Then
                ArrayCmp = -1
                Exit Function
            Else
                UBoundNum1 = UBoundNum1 - 1
            End If
        Loop
    End If

End Function

Public Sub AddArray(Num1() As Integer, ByVal UBoundNum1 As Long, Num2() As Integer, ByVal UBoundNum2 As Long, Sum() As Integer, UBoundSum As Long)

    Dim Carry As Long
    Dim tmpSum As Long
    Dim i As Long

    If UBoundNum1 > UBoundNum2 Then
        UBoundSum = UBoundNum1 + 1
        ReDim Sum(UBoundSum)
        For i = 0 To UBoundNum2
            tmpSum = Num1(i) + Num2(i) + Carry
            If tmpSum >= 10000 Then
                Carry = 1
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Sum(i) = tmpSum
        Next i
        Do While Carry
            If i > UBoundNum1 Then Exit Do
            tmpSum = Num1(i) + Carry
            If tmpSum >= 10000 Then
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Sum(i) = tmpSum
            i = i + 1
        Loop
        For i = i To UBoundNum1
            Sum(i) = Num1(i)
        Next i
    Else
        UBoundSum = UBoundNum2 + 1
        ReDim Sum(UBoundSum)
        For i = 0 To UBoundNum1
            tmpSum = Num1(i) + Num2(i) + Carry
            If tmpSum >= 10000 Then
                Carry = 1
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Sum(i) = tmpSum
        Next i
        Do While Carry
            If i > UBoundNum2 Then Exit Do
            tmpSum = Num2(i) + Carry
            If tmpSum >= 10000 Then
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Sum(i) = tmpSum
            i = i + 1
        Loop
        For i = i To UBoundNum2
            Sum(i) = Num2(i)
        Next i
    End If
    Sum(i) = Carry

    For i = UBoundSum To 0 Step -1
        If Sum(i) <> 0 Then Exit For
    Next i
    If i = UBoundSum Then Exit Sub
    If i < 0 Then i = 0
    ReDim Preserve Sum(i)
    UBoundSum = i

End Sub

Public Sub AddArray2(Num1Sum() As Integer, UBoundNum1Sum As Long, Num2() As Integer, ByVal UBoundNum2 As Long)

    Dim Carry As Long
    Dim tmpSum As Long
    Dim i As Long
    Dim UBoundNum1 As Long

    UBoundNum1 = UBoundNum1Sum
    If UBoundNum1 > UBoundNum2 Then
        UBoundNum1Sum = UBoundNum1 + 1
        ReDim Preserve Num1Sum(UBoundNum1Sum)
        For i = 0 To UBoundNum2
            tmpSum = Num1Sum(i) + Num2(i) + Carry
            If tmpSum >= 10000 Then
                Carry = 1
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Num1Sum(i) = tmpSum
        Next i
        Do While Carry
            If i > UBoundNum1 Then Exit Do
            tmpSum = Num1Sum(i) + Carry
            If tmpSum >= 10000 Then
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Num1Sum(i) = tmpSum
            i = i + 1
        Loop
    Else
        UBoundNum1Sum = UBoundNum2 + 1
        ReDim Preserve Num1Sum(UBoundNum1Sum)
        For i = 0 To UBoundNum1
            tmpSum = Num1Sum(i) + Num2(i) + Carry
            If tmpSum >= 10000 Then
                Carry = 1
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Num1Sum(i) = tmpSum
        Next i
        Do While Carry
            If i > UBoundNum2 Then Exit Do
            tmpSum = Num2(i) + Carry
            If tmpSum >= 10000 Then
                tmpSum = tmpSum - 10000
            Else
                Carry = 0
            End If
            Num1Sum(i) = tmpSum
            i = i + 1
        Loop
        For i = i To UBoundNum2
            Num1Sum(i) = Num2(i)
        Next i
    End If
    Num1Sum(UBoundNum1Sum) = Carry

    For i = UBoundNum1Sum To 0 Step -1
        If Num1Sum(i) <> 0 Then Exit For
    Next i
    If i = UBoundNum1Sum Then Exit Sub
    If i < 0 Then i = 0
    ReDim Preserve Num1Sum(i)
    UBoundNum1Sum = i

End Sub

Public Sub SubArray(Num1() As Integer, ByVal UBoundNum1 As Long, Num2() As Integer, ByVal UBoundNum2 As Long, Diff() As Integer, UBoundDiff As Long)

    Dim Carry As Long
    Dim tmpDiff As Long
    Dim i As Long

    UBoundDiff = UBoundNum1
    ReDim Diff(UBoundDiff)
    For i = 0 To UBoundNum2
        tmpDiff = Num1(i) - Num2(i) - Carry
        If tmpDiff < 0 Then
            tmpDiff = tmpDiff + 10000
            Carry = 1
        Else
            Carry = 0
        End If
        Diff(i) = tmpDiff
    Next i
    Do While Carry
        If i > UBoundNum1 Then Exit Do
        tmpDiff = Num1(i) - Carry
        If tmpDiff < 0 Then
            tmpDiff = tmpDiff + 10000
        Else
            Carry = 0
        End If
        Diff(i) = tmpDiff
        i = i + 1
    Loop
    For i = i To UBoundNum1
        Diff(i) = Num1(i)
    Next i

    For i = UBoundDiff To 0 Step -1
        If Diff(i) <> 0 Then Exit For
    Next i
    If i = UBoundDiff Then Exit Sub
    If i < 0 Then i = 0
    ReDim Preserve Diff(i)
    UBoundDiff = i

End Sub

Public Sub SubArray2(Num1Diff() As Integer, UBoundNum1Diff As Long, Num2() As Integer, ByVal UBoundNum2 As Long)

    Dim Carry As Long
    Dim tmpDiff As Long
    Dim i As Long

    For i = 0 To UBoundNum2
        tmpDiff = Num1Diff(i) - Num2(i) - Carry
        If tmpDiff < 0 Then
            tmpDiff = tmpDiff + 10000
            Carry = 1
        Else
            Carry = 0
        End If
        Num1Diff(i) = tmpDiff
    Next i
    Do While Carry
        If i > UBoundNum1Diff Then Exit Do
        tmpDiff = Num1Diff(i) - Carry
        If tmpDiff < 0 Then
            tmpDiff = tmpDiff + 10000
        Else
            Carry = 0
        End If
        Num1Diff(i) = tmpDiff
        i = i + 1
    Loop

    For i = UBoundNum1Diff To 0 Step -1
        If Num1Diff(i) <> 0 Then Exit For
    Next i
    If i = UBoundNum1Diff Then Exit Sub
    If i < 0 Then i = 0
    ReDim Preserve Num1Diff(i)
    UBoundNum1Diff = i

End Sub

Public Sub ArrayMultiply(Num1() As Integer, ByVal UBoundNum1 As Long, Num2() As Integer, ByVal UBoundNum2 As Long, Ans() As Integer, UBoundAns As Long)

    Dim i1 As Long
    Dim i2 As Long
    Dim ia As Long
    Dim Tmp As Long
    Dim Carry  As Long

    UBoundAns = UBoundNum1 + UBoundNum2 + 1
    ReDim Ans(UBoundAns)

    For i1 = 0 To UBoundNum1
        ia = i1
        Carry = 0
        For i2 = 0 To UBoundNum2
            Tmp = CLng(Num2(i2)) * Num1(i1) + Carry + Ans(ia)
            Ans(ia) = Tmp Mod 10000
            Carry = Tmp \ 10000
            ia = ia + 1
        Next i2
        Ans(ia) = Carry
    Next i1

    For i1 = UBoundAns To 0 Step -1
        If Ans(i1) <> 0 Then Exit For
    Next i1
    If i1 = UBoundAns Then Exit Sub
    If i1 < 0 Then i1 = 0
    ReDim Preserve Ans(i1)
    UBoundAns = i1

End Sub

Public Function ArrayDivide(Dvdnd() As Integer, ByVal UBoundDvdnd As Long, Divisor() As Integer, ByVal UBoundDivisor As Long, Quo() As Integer, UBoundQuo As Long, Rmndr() As Integer, UBoundRmndr As Long) As Boolean

    Dim tmpQ As Integer
    Dim tmpDivisor As Integer
    Dim tmpArr() As Integer
    Dim tmpAns As Long
    Dim Carry As Long
    Dim i As Long
    Dim j As Long

    UBoundQuo = UBoundDvdnd - UBoundDivisor
    If UBoundQuo < 0 Then
        UBoundQuo = 0
        ReDim Quo(0)
        CopyArray Dvdnd, UBoundDvdnd, Rmndr, UBoundRmndr
        Exit Function
    End If

    ReDim Quo(UBoundQuo)
    UBoundRmndr = UBoundDivisor + 1
    ReDim Rmndr(UBoundRmndr)
    ReDim tmpArr(UBoundRmndr)
    For i = UBoundDivisor To 0 Step -1
        Rmndr(i) = Dvdnd(i + (UBoundDvdnd - UBoundDivisor))
    Next i
    tmpDivisor = Divisor(UBoundDivisor)

    Do While 1
        ''''''''''' tmpAns and Carry used as temporary variables
        If UBoundDivisor = 0 Then
            tmpAns = 0
            Carry = 0
        Else
            tmpAns = Rmndr(UBoundDivisor - 1)
            Carry = Divisor(UBoundDivisor - 1)
        End If
        tmpQ = ((CDbl(Rmndr(UBoundRmndr)) * 10000 + CDbl(Rmndr(UBoundDivisor))) * 10000 + tmpAns) / (CDbl(tmpDivisor) * 10000 + Carry)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If tmpQ > 9999 Then tmpQ = 9999
        Carry = 0
        For i = 0 To UBoundDivisor
            tmpAns = CLng(Divisor(i)) * tmpQ + Carry
            Carry = tmpAns \ 10000
            tmpArr(i) = tmpAns Mod 10000
        Next i
        tmpArr(i) = Carry
        For i = UBoundRmndr To 0 Step -1
            If Rmndr(i) > tmpArr(i) Then
                Exit For
            ElseIf Rmndr(i) < tmpArr(i) Then
                Carry = 0
                For j = 0 To UBoundDivisor
                    tmpAns = tmpArr(j) - Divisor(j) - Carry
                    If tmpAns < 0 Then
                        tmpAns = tmpAns + 10000
                        Carry = 1
                    Else
                        Carry = 0
                    End If
                    tmpArr(j) = tmpAns
                Next j
                If Carry Then tmpArr(j) = tmpArr(j) - Carry
                tmpQ = tmpQ - 1
                i = UBoundRmndr + 1
            End If
        Next i
        Quo(UBoundDvdnd - UBoundDivisor) = tmpQ
        Carry = 0
        For i = 0 To UBoundRmndr
            tmpAns = Rmndr(i) - tmpArr(i) - Carry
            If tmpAns < 0 Then
                tmpAns = tmpAns + 10000
                Carry = 1
            Else
                Carry = 0
            End If
            Rmndr(i) = tmpAns
        Next i
        UBoundDvdnd = UBoundDvdnd - 1
        If UBoundDvdnd < UBoundDivisor Then Exit Do
        For i = UBoundDivisor To 0 Step -1
            Rmndr(i + 1) = Rmndr(i)
        Next i
        Rmndr(0) = Dvdnd(UBoundDvdnd - UBoundDivisor)
    Loop

    For i = UBoundRmndr To 0 Step -1
        If Rmndr(i) <> 0 Then Exit For
    Next i
    If i < 0 Then i = 0
    If i <> UBoundRmndr Then ReDim Preserve Rmndr(i)
    UBoundRmndr = i
    If UBoundRmndr = 0 And Rmndr(0) = 0 Then ArrayDivide = True

    For i = UBoundQuo To 0 Step -1
        If Quo(i) <> 0 Then Exit For
    Next i
    If i = UBoundQuo Then Exit Function
    If i < 0 Then i = 0
    ReDim Preserve Quo(i)
    UBoundQuo = i

End Function

Public Sub Sqrt(Num() As Integer, ByVal UBoundNum As Long, Ans() As Integer, UBoundAns As Long)

    Dim tmpDiv() As Integer
    Dim UBDiv As Long
    Dim tmpRmndr() As Integer
    Dim UBRmndr As Long
    Dim tmpArr() As Integer
    Dim tmpQ As Integer
    Dim tmpAns As Long
    Dim Carry As Long
    Dim i As Long
    Dim j As Long

    UBoundAns = UBoundNum \ 2
    ReDim Ans(UBoundAns)
    UBRmndr = 1
    ReDim tmpRmndr(1)
    ReDim tmpDiv(0)
    If UBoundNum And 1 Then
        tmpRmndr(1) = Num(UBoundNum)
        tmpRmndr(0) = Num(UBoundNum - 1)
        tmpDiv(0) = Fix(Sqr(CLng(Num(UBoundNum)) * 10000 + Num(UBoundNum - 1)))
    Else
        tmpRmndr(0) = Num(UBoundNum)
        tmpDiv(0) = Fix(Sqr(Num(UBoundNum)))
        UBoundNum = UBoundNum + 1
    End If

    Do While 1
        ReDim tmpArr(UBRmndr)
        ''''''''''' tmpAns and Carry used as temporary variables
        If UBDiv = 0 Then
            tmpAns = 0
            Carry = 0
        Else
            tmpAns = tmpRmndr(UBDiv - 1)
            Carry = tmpDiv(UBDiv - 1)
        End If
        tmpQ = ((CDbl(tmpRmndr(UBRmndr)) * 10000 + CDbl(tmpRmndr(UBDiv))) * 10000 + tmpAns) / (CDbl(tmpDiv(UBDiv)) * 10000 + Carry)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If tmpQ > 9999 Then tmpQ = 9999

        tmpDiv(0) = tmpQ
        Carry = 0
        For i = 0 To UBDiv
            tmpAns = CLng(tmpDiv(i)) * tmpQ + Carry
            Carry = tmpAns \ 10000
            tmpArr(i) = tmpAns Mod 10000
        Next i
        tmpArr(i) = Carry
        For i = UBRmndr To 0 Step -1
            If tmpRmndr(i) > tmpArr(i) Then
                Exit For
            ElseIf tmpRmndr(i) < tmpArr(i) Then
                tmpQ = tmpQ - 1
                tmpAns = tmpArr(0) - tmpDiv(0) - tmpQ
                Carry = 0
                Do While tmpAns < 0
                    tmpAns = tmpAns + 10000
                    Carry = Carry + 1
                Loop
                tmpArr(0) = tmpAns
                For j = 1 To UBDiv
                    tmpAns = tmpArr(j) - tmpDiv(j) - Carry
                    If tmpAns < 0 Then
                        tmpAns = tmpAns + 10000
                        Carry = 1
                    Else
                        Carry = 0
                    End If
                    tmpArr(j) = tmpAns
                Next j
                If Carry Then tmpArr(j) = tmpArr(j) - Carry
                tmpDiv(0) = tmpQ
                i = UBRmndr + 1
            End If
        Next i

        Ans(UBoundNum \ 2) = tmpQ
        UBoundNum = UBoundNum - 2
        If UBoundNum < 0 Then Exit Do

        UBDiv = UBoundAns - (UBoundNum \ 2)
        If Ans(UBoundAns) > 4999 Then UBDiv = UBDiv + 1
        ReDim tmpDiv(UBDiv)
        Carry = 0
        j = 1
        For i = UBoundNum \ 2 + 1 To UBoundAns
            tmpAns = Ans(i) * 2 + Carry
            If tmpAns >= 10000 Then
                tmpAns = tmpAns - 10000
                Carry = 1
            Else
                Carry = 0
            End If
            tmpDiv(j) = tmpAns
            j = j + 1
        Next i
        If Carry Then tmpDiv(j) = Carry

        Carry = 0
        For i = 0 To UBRmndr
            tmpAns = tmpRmndr(i) - tmpArr(i) - Carry
            If tmpAns < 0 Then
                tmpAns = tmpAns + 10000
                Carry = 1
            Else
                Carry = 0
            End If
            tmpRmndr(i) = tmpAns
        Next i
        UBRmndr = UBDiv + 1
        ReDim Preserve tmpRmndr(UBRmndr)
        For i = UBRmndr To 2 Step -1
            tmpRmndr(i) = tmpRmndr(i - 2)
        Next i
        tmpRmndr(1) = Num(UBoundNum)
        tmpRmndr(0) = Num(UBoundNum - 1)
    Loop

End Sub

Public Sub PowerOf2(Power As Long, Ans() As Integer, UBoundAns As Long)

    Dim i As Long
    Dim Mask As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim tmpPow() As Integer
    Dim UBtmpPow As Long

    If Not DidInit Then
        Init
    End If

    ReDim Ans(0)
    UBoundAns = 0
    If Power < 14 Then
        Ans(0) = BitMask(Power)
        Exit Sub
    End If

    ReDim tmpPow(0)
    tmpPow(0) = 256
    Mask = Power And 7
    Ans(0) = BitMask(Mask)
    i = 3
    Mask = Power - Mask

    Do While 1
        If Mask And BitMask(i) Then
            ArrayMultiply tmpPow, UBtmpPow, Ans, UBoundAns, tmpAns, UBtmpAns
            CopyArray tmpAns, UBtmpAns, Ans, UBoundAns
        End If
        Mask = Mask And Not BitMask(i)
        If Mask = 0 Then Exit Do
        ArrayMultiply tmpPow, UBtmpPow, tmpPow, UBtmpPow, tmpAns, UBtmpAns
        CopyArray tmpAns, UBtmpAns, tmpPow, UBtmpPow
        i = i + 1
    Loop

End Sub

Public Sub ArrayPower(Num() As Integer, ByVal UBNum As Long, Power() As Integer, ByVal UBPower As Long, Ans() As Integer, UBAns As Long)

    Dim i As Long
    Dim Mask() As Integer
    Dim UBMask As Long
    Dim LenMask As Long
    Dim tmpAns() As Integer
    Dim UBtmpAns As Long
    Dim tmpPow() As Integer
    Dim UBtmpPow As Long

    ReDim Ans(0)
    UBAns = 0
    Ans(0) = 1

    CopyArray Num, UBNum, tmpPow, UBtmpPow
    ArrayToBin Power, UBPower, Mask, UBMask

    LenMask = LenBin(Mask, UBMask)

    For i = 1 To LenMask
        If IsNthBitSet(i, Mask, UBMask) Then
            ArrayMultiply tmpPow, UBtmpPow, Ans, UBAns, tmpAns, UBtmpAns
            CopyArray tmpAns, UBtmpAns, Ans, UBAns
        End If
        If i = LenMask Then Exit For
        ArrayMultiply tmpPow, UBtmpPow, tmpPow, UBtmpPow, tmpAns, UBtmpAns
        CopyArray tmpAns, UBtmpAns, tmpPow, UBtmpPow
    Next i

End Sub

Public Function PrimeFactorsOfArray(Num() As Integer, UBNum As Long) As String

    Dim tmpArr() As Integer
    Dim UBtmpArr As Long
    Dim Div() As Integer
    Dim UBDiv As Long
    Dim Q() As Integer
    Dim UBQ As Long
    Dim R() As Integer
    Dim UBR As Long

    CopyArray Num, UBNum, tmpArr, UBtmpArr
    If Not IsArrayPrime2(tmpArr, UBtmpArr, Div, UBDiv) Then
        Do
            PrimeFactorsOfArray = PrimeFactorsOfArray & ArrayToString(Div, UBDiv) & "*"
            ArrayDivide tmpArr, UBtmpArr, Div, UBDiv, Q, UBQ, R, UBR
            CopyArray Q, UBQ, tmpArr, UBtmpArr
        Loop While Not IsArrayPrime2(tmpArr, UBtmpArr, Div, UBDiv)
        PrimeFactorsOfArray = PrimeFactorsOfArray & ArrayToString(tmpArr, UBtmpArr)
    End If

End Function

Public Function LenBin(Bin() As Integer, ByVal UBBin As Long) As Long

    Dim Bits As Integer

    Bits = Bin(UBBin)
    LenBin = UBBin * 4
    If Bits >= 1000 Then
        LenBin = LenBin + 4
    ElseIf Bits >= 100 Then
        LenBin = LenBin + 3
    ElseIf Bits >= 10 Then
        LenBin = LenBin + 2
    Else
        LenBin = LenBin + 1
    End If

End Function

Public Function IsNthBitSet(ByVal N As Long, Bin() As Integer, ByVal UBBin As Long) As Boolean

    Dim i As Integer
    Dim Bits As Integer

    N = N - 1
    Bits = Bin(N \ 4)
    For i = 1 To N Mod 4
        Bits = Bits \ 10
    Next i
    IsNthBitSet = Bits Mod 10

End Function

Public Function HexBitsToDec(ByVal Bits As Integer) As Integer

    Select Case Bits
    Case 0
        HexBitsToDec = 0
    Case 1
        HexBitsToDec = 1
    Case 10
        HexBitsToDec = 2
    Case 11
        HexBitsToDec = 3
    Case 100
        HexBitsToDec = 4
    Case 101
        HexBitsToDec = 5
    Case 110
        HexBitsToDec = 6
    Case 111
        HexBitsToDec = 7
    Case 1000
        HexBitsToDec = 8
    Case 1001
        HexBitsToDec = 9
    Case 1010
        HexBitsToDec = 10
    Case 1011
        HexBitsToDec = 11
    Case 1100
        HexBitsToDec = 12
    Case 1101
        HexBitsToDec = 13
    Case 1110
        HexBitsToDec = 14
    Case 1111
        HexBitsToDec = 15
    End Select

End Function

Public Function HexBitsToHexDigit(ByVal Bits As Integer) As String

    Select Case Bits
    Case 0
        HexBitsToHexDigit = "0"
    Case 1
        HexBitsToHexDigit = "1"
    Case 10
        HexBitsToHexDigit = "2"
    Case 11
        HexBitsToHexDigit = "3"
    Case 100
        HexBitsToHexDigit = "4"
    Case 101
        HexBitsToHexDigit = "5"
    Case 110
        HexBitsToHexDigit = "6"
    Case 111
        HexBitsToHexDigit = "7"
    Case 1000
        HexBitsToHexDigit = "8"
    Case 1001
        HexBitsToHexDigit = "9"
    Case 1010
        HexBitsToHexDigit = "a"
    Case 1011
        HexBitsToHexDigit = "b"
    Case 1100
        HexBitsToHexDigit = "c"
    Case 1101
        HexBitsToHexDigit = "d"
    Case 1110
        HexBitsToHexDigit = "e"
    Case 1111
        HexBitsToHexDigit = "f"
    End Select

End Function

Private Sub Init()

    HexBits(0) = 0
    HexBits(1) = 1
    HexBits(2) = 10
    HexBits(3) = 11
    HexBits(4) = 100
    HexBits(5) = 101
    HexBits(6) = 110
    HexBits(7) = 111
    HexBits(8) = 1000
    HexBits(9) = 1001
    HexBits(10) = 1010
    HexBits(11) = 1011
    HexBits(12) = 1100
    HexBits(13) = 1101
    HexBits(14) = 1110
    HexBits(15) = 1111

    BitMask(0) = &H1
    BitMask(1) = &H2
    BitMask(2) = &H4
    BitMask(3) = &H8
    BitMask(4) = &H10
    BitMask(5) = &H20
    BitMask(6) = &H40
    BitMask(7) = &H80
    BitMask(8) = &H100
    BitMask(9) = &H200
    BitMask(10) = &H400
    BitMask(11) = &H800
    BitMask(12) = &H1000
    BitMask(13) = &H2000
    BitMask(14) = &H4000
    BitMask(15) = 32768                    'DO NOT use &H8000 here cause VB makes it &HFFFF8000 internally while putting it in a Long, extending the SIGN BIT of the INT &H8000
    BitMask(16) = &H10000
    BitMask(17) = &H20000
    BitMask(18) = &H40000
    BitMask(19) = &H80000
    BitMask(20) = &H100000
    BitMask(21) = &H200000
    BitMask(22) = &H400000
    BitMask(23) = &H800000
    BitMask(24) = &H1000000
    BitMask(25) = &H2000000
    BitMask(26) = &H4000000
    BitMask(27) = &H8000000
    BitMask(28) = &H10000000
    BitMask(29) = &H20000000
    BitMask(30) = &H40000000
    BitMask(31) = &H80000000

    DidInit = True

End Sub
