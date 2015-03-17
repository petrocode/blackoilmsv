Module Mathematics

    Function Linear_Interpolation_3D(ByVal base As Integer, ByVal target As Integer, ByVal relative(,,) As Decimal, ByVal sw_value As Decimal, ByVal sg_value As Decimal) As Decimal
        '       Dim m As Decimal
        '       Dim sw_1, sw_2, out_1, out_2 As Decimal

        '       For i As Integer = 0 To relative.GetLength(0) - 1
        '           If relative Then
        '       Next
        'if ((relative.sw[i]>=sw_value && relative.sw[i+1]<=sw_value ) || (relative.sw[i]<=sw_value && relative.sw[i+1]>=sw_value ))   {
        '	   sw_1=base,target,relative.sw[i];
        '	   sw_2=base,target,relative.sw[i+1];
        '	   out_1=linear_interpolation_2D(base,target,relative.data[i],sg_value);
        '	   out_2=linear_interpolation_2D(base,target,relative.data[i+1],sg_value);
        '	   m = (out_2-out_1)/(sw_2-sw_1);
        '	   return (m*(sw_value-sw_1) + out_1);
    End Function


    Function Slope(ByVal y2 As Decimal, ByVal y1 As Decimal, ByVal x2 As Decimal, ByVal x1 As Decimal) As Decimal
        If x2 <> x1 Then
            Return ((y2 - y1) / (x2 - x1))
        End If
    End Function


    Function Calculate_Mse(ByVal x(,) As Decimal, ByVal A(,) As Decimal, ByVal d(,) As Decimal, ByVal x_d(,) As Decimal) As Decimal
        Multiply_Matrix(A, x, x_d)
        Dim MSE_Total As Decimal
        For j As Integer = 0 To x.GetLength(1) - 1
            MSE_Total += (d(0, j) - x_d(0, j)) ^ 2
        Next
        Return MSE_Total / x.GetLength(1)
    End Function


    Sub Multiply_Matrix(ByVal m1(,) As Decimal, ByVal m2(,) As Decimal, ByRef out(,) As Decimal)

        Dim sum As Decimal
        For j1 As Integer = 0 To m1.GetLength(1) - 1
            For j2 As Integer = 0 To m2.GetLength(0) - 1
                sum = 0
                For i As Integer = 0 To m1.GetLength(0) - 1
                    sum += m1(i, j1) * m2(j2, i)
                Next
                out(j2, j1) = sum
            Next
        Next
    End Sub


    Function Dot_Product(ByVal m1(,) As Decimal, ByVal m2(,) As Decimal) As Decimal
        Dim sum As Decimal
        For i As Integer = 0 To m1.GetLength(1) - 1
            sum += m1(0, i) * m2(0, i)
        Next
        Return sum
    End Function


    Sub Add_Matrix(ByVal m1(,) As Decimal, ByVal m2(,) As Decimal, ByRef out(,) As Decimal, Optional ByVal Negative As Boolean = False)
        For i As Integer = 0 To m1.GetLength(0) - 1
            For j As Integer = 0 To m1.GetLength(1) - 1
                If Negative Then
                    out(i, j) = m1(i, j) - m2(i, j)
                Else
                    out(i, j) = m1(i, j) + m2(i, j)
                End If
            Next
        Next
    End Sub

   
    Sub Equal_Matrix(ByVal m1(,) As Decimal, ByRef m2(,) As Decimal)
        For i As Integer = 0 To m1.GetLength(0) - 1
            For j As Integer = 0 To m1.GetLength(1) - 1
                m2(i, j) = m1(i, j)
            Next
        Next
    End Sub

    Function Solve_LSOR(ByVal A(,) As Decimal, ByVal x(,) As Decimal, ByVal d(,) As Decimal, ByVal K As Integer, ByVal mse As Decimal) As Boolean
        Dim sum As Decimal
        Dim cur_mse As Decimal = mse
        Dim n As Integer = 0
        Dim x_d(0, x.GetLength(1) - 1) As Decimal

        Do While n < K And cur_mse >= mse

            For i As Integer = 0 To x.GetLength(1) - 1
                sum = 0
                For j As Integer = 0 To x.GetLength(1) - 1
                    If j <> i Then
                        sum += x(0, j) * A(j, i)
                    End If
                Next
                x(0, i) = (d(0, i) - sum) / (A(i, i))
            Next

            cur_mse = Calculate_Mse(x, A, d, x_d)

            n += 1
        Loop



    End Function
    Function L_Min(ByVal Ng As Integer, ByVal DR As Decimal, ByVal L_Core As Decimal) As Decimal
        Return (1 - DR / 100) * (L_Core / Ng)
    End Function
    Function SumOfArray(ByVal Arr() As Decimal) As Decimal
        Dim out As Decimal
        For i As Integer = 0 To SizeOf(Arr) - 1
            out += Arr(i)
        Next
        Return out
    End Function
    Function ReverseArray(ByVal Arr() As Decimal) As Decimal()
        Dim out(SizeOf(Arr) - 1) As Decimal
        For i As Integer = 0 To SizeOf(Arr) - 1
            out(i) = 1 / Arr(i)
        Next
        Return out
    End Function
    Function ProductArrays(ByVal Arr1() As Decimal, ByVal Arr2() As Decimal) As Decimal()
        Dim out(SizeOf(Arr1) - 1) As Decimal
        For i As Integer = 0 To SizeOf(Arr1) - 1
            out(i) = Arr1(i) * Arr2(i)
        Next
        Return out
    End Function
    Function Normalize(ByVal Values() As Decimal) As Decimal()
        Dim out(SizeOf(Values) - 1) As Decimal
        Dim max As Decimal = Values(0)
        Dim min As Decimal = max
        For i As Integer = 0 To SizeOf(Values) - 1
            If Values(i) > max Then max = Values(i)
            If Values(i) < min Then min = Values(i)
        Next
        For i As Integer = 0 To SizeOf(Values) - 1
            out(i) = (Values(i) - min) / (max - min)
        Next
        Return out
    End Function

    Function Solve(ByVal A(,) As Decimal, ByVal x(,) As Decimal, ByVal d(,) As Decimal, ByVal k As Integer, ByVal mse As Decimal, ByVal Method As Equation_Solution_Method) As Boolean
        Select Case Method
            Case Equation_Solution_Method.LSOR
                Solve_LSOR(A, x, d, k, mse)
            Case Equation_Solution_Method.Orhtomin
                Solve_Orthomin(A, x, d, k, mse)
        End Select

    End Function
    Function Solve_Orthomin(ByVal _A(,) As Decimal, ByRef _x(,) As Decimal, ByVal _d(,) As Decimal, ByVal k As Integer, ByVal mse As Decimal) As Boolean
        Dim A(_A.GetLength(0) - 1, _A.GetLength(1) - 1) As Decimal
        Dim X(_x.GetLength(0) - 1, _x.GetLength(1) - 1) As Decimal
        Dim d(_d.GetLength(0) - 1, _d.GetLength(1) - 1) As Decimal
        For i1 As Integer = 0 To _A.GetLength(0) - 1
            For i2 As Integer = 0 To _A.GetLength(1) - 1
                A(i1, i2) = CDec(_A(i1, i2))
            Next
            X(0, i1) = CDec(_x(0, i1))
            d(0, i1) = CDec(_d(0, i1))
        Next


        Dim restart As Integer = X.GetLength(1)
        Dim n As Integer


        Dim answers(k - 1, X.GetLength(0) - 1, X.GetLength(1) - 1) As Decimal
        Dim MSEs(k - 1) As Decimal
        Dim sum_bp(0, A.GetLength(1) - 1) As Decimal
        Dim x_d(0, A.GetLength(1) - 1) As Decimal
        Dim ar(0, A.GetLength(1) - 1) As Decimal
        Dim ag(0, A.GetLength(1) - 1) As Decimal
        Dim ap(0, A.GetLength(1) - 1) As Decimal
        Dim A_x(0, X.GetLength(1) - 1) As Decimal

        Dim g(restart - 1)(,) As Decimal
        Dim delta(restart - 1)(,) As Decimal
        Dim b(restart - 1)(,) As Decimal
        Dim bp(restart - 1)(,) As Decimal
        Dim p(restart - 1)(,) As Decimal

        For i As Integer = 0 To restart - 1
            ReDim g(i)(0, X.GetLength(1) - 1)
            ReDim delta(i)(0, 0)
            ReDim b(i)(0, 0)
            ReDim bp(i)(0, X.GetLength(1) - 1)
            ReDim p(i)(0, X.GetLength(1) - 1)
        Next

        Dim aa(0, 0) As Decimal
        Dim r(0, X.GetLength(1) - 1) As Decimal

        initialize_OM(A, X, d, r, p, A_x)

        Dim Current_MSE As Decimal = mse
        Dim Last_MSE As Decimal = 1000000000000
        Dim j As Integer

        Do
            perform_OM(A, X, d, r, p, aa, ap, ag, ar, g, delta, b, bp, sum_bp, j)

            Current_MSE = Calculate_Mse(X, A, d, x_d)

            CopyAnswers(n, answers, X)
            MSEs(n) = Current_MSE

            If j = restart - 1 Then
                Equal_Matrix(p(j), p(0))
                j = 0
            Else
                j += 1
            End If

            Last_MSE = Current_MSE
            n += 1
        Loop While Current_MSE > mse And n < k
        If k = n Then
            ' Stop
        End If
        ReturnBestOutput(answers, _x, MSEs, n)

        Return True

    End Function
    Sub ReturnBestOutput(ByVal answers(,,) As Decimal, ByRef x(,) As Decimal, ByVal MSEs() As Decimal, ByVal count As Integer)
        Dim curMSE As Decimal = MSEs(0)
        Dim outI As Integer
        For i As Integer = 0 To count - 1
            If MSEs(i) < curMSE And Not Double.IsNaN(MSEs(i)) Then
                curMSE = MSEs(i)
                outI = i
            End If
        Next

        For i As Integer = 0 To x.GetLength(0) - 1
            For j As Integer = 0 To x.GetLength(1) - 1
                x(i, j) = answers(outI, i, j)
            Next
        Next

    End Sub


    
    Sub CopyAnswers(ByVal n As Integer, ByRef answers(,,) As Decimal, ByVal x(,) As Decimal)
        For i As Integer = 0 To answers.GetLength(1) - 1
            For j As Integer = 0 To answers.GetLength(2) - 1
                answers(n, i, j) = x(i, j)
            Next
        Next
    End Sub


    Sub initialize_OM(ByVal A(,) As Decimal, ByVal x(,) As Decimal, ByVal d(,) As Decimal, ByRef r(,) As Decimal, ByRef p()(,) As Decimal, ByRef A_x(,) As Decimal)
        Multiply_Matrix(A, x, A_x)
        Add_Matrix(d, A_x, r, True)
        Equal_Matrix(r, p(0))
    End Sub

    Sub perform_OM(ByRef A(,) As Decimal, ByRef x(,) As Decimal, ByRef d(,) As Decimal, ByRef r(,) As Decimal, ByRef p()(,) As Decimal, ByRef aa(,) As Decimal, ByRef ap(,) As Decimal, ByRef ag(,) As Decimal, ByRef ar(,) As Decimal, ByRef g()(,) As Decimal, ByRef delta()(,) As Decimal, ByRef b()(,) As Decimal, ByRef bp()(,) As Decimal, ByRef sum_bp(,) As Decimal, ByVal stp As Integer)
        Multiply_Matrix(A, p(stp), g(stp))
        delta(stp)(0, 0) = Dot_Product(g(stp), g(stp))
        aa(0, 0) = Dot_Product(r, g(stp)) / delta(stp)(0, 0)
        Multiply_Matrix(p(stp), aa, ap)
        Add_Matrix(x, ap, x, False)
        Multiply_Matrix(g(stp), aa, ag)
        Add_Matrix(r, ag, r, True)
        Multiply_Matrix(A, r, ar)

        For i As Integer = 0 To sum_bp.GetLength(0) - 1
            For j As Integer = 0 To sum_bp.GetLength(1) - 1
                sum_bp(i, j) = 0
            Next
        Next

        For i As Integer = 0 To stp
            b(i)(0, 0) = -Dot_Product(ar, g(i)) / delta(i)(0, 0)
            Multiply_Matrix(p(i), b(i), bp(i))
            Add_Matrix(sum_bp, bp(i), sum_bp, False)
        Next
        If stp + 1 < SizeOf(p) Then
            Add_Matrix(r, sum_bp, p(stp + 1), False)
        End If



    End Sub

    Function Interpolate(ByVal Values(,) As Decimal, ByVal value As Decimal) As Decimal
        For i As Integer = 0 To Values.GetLength(1) - 2
            If value = Values(0, i) Then Return Values(1, i)
            If Values(0, i) < value And Values(0, i + 1) > value Then
                Dim m As Decimal = (Values(1, i + 1) - Values(1, i)) / (Values(0, i + 1) - Values(0, i))
                Return m * (value - Values(0, i)) + Values(1, i)
            End If
        Next
    End Function


    Function Common_Area(ByVal side1 As Side, ByVal side2 As Side) As Decimal
        Dim x, y As Decimal
        Dim xmax1, xmin1, xmax2, xmin2 As Decimal
        Dim ymax1, ymin1, ymax2, ymin2 As Decimal

        xmax1 = side1.xCoord + (side1.xLen) / 2
        xmin1 = side1.xCoord - (side1.xLen) / 2
        xmax2 = side2.xCoord + (side2.xLen) / 2
        xmin2 = side2.xCoord - (side2.xLen) / 2

        ymax1 = side1.yCoord + (side1.yLen) / 2
        ymin1 = side1.yCoord - (side1.yLen) / 2
        ymax2 = side2.yCoord + (side2.yLen) / 2
        ymin2 = side2.yCoord - (side2.yLen) / 2

        Return Common_Length(xmax1, xmin1, xmax2, xmin2) * Common_Length(ymax1, ymin1, ymax2, ymin2)

    End Function

    Function Common_Length(ByVal max1 As Decimal, ByVal min1 As Decimal, ByVal max2 As Decimal, ByVal min2 As Decimal) As Decimal

        If ((max1 >= max2 And min1 >= max2) Or (max1 <= min2 And min1 <= min2)) Then Return 0

        If (max1 > max2 And min1 > min2) Then Return max2 - min1
        If (max1 >= max2 And min1 <= min2) Then Return max2 - min2
        If (max1 < max2 And min1 < min2) Then Return max1 - min2
        If (max2 > max1 And min2 > min1) Then Return max1 - min2
        If (max2 >= max1 And min2 <= min1) Then Return max1 - min1
        If (max2 < max1 And min2 < min1) Then Return max2 - min1

    End Function

    Function Linear_Interpolation_2D(ByVal Value As Decimal, ByVal data(,) As Decimal, ByVal Base As Integer, ByVal ParamArray Columns() As Integer) As Decimal()
        Dim out(Columns.Length - 1) As Decimal
        For i As Integer = 0 To data.GetLength(1) - 2
            If data(Base, i) <= Value And data(Base, i + 1) >= Value Then
                For n As Integer = 0 To out.Length - 1
                    Dim m As Decimal = (data(Columns(n), i + 1) - data(Columns(n), i)) / (data(Base, i + 1) - data(Base, i))
                    out(n) = data(Columns(n), i) + m * (Value - data(Base, i))
                Next
                Return out
            End If
        Next
    End Function

    Function Average_Permeability(ByVal k1 As Decimal, ByVal k2 As Decimal, ByVal l1 As Decimal, ByVal l2 As Decimal) As Decimal
        If k1 = 0 Or k2 = 0 Then Return 0
        Return (l1 + l2) / ((l1 / k1) + (l2 / k2))
    End Function



End Module
