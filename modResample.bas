Attribute VB_Name = "modResample"
Option Explicit

Enum eTypes
    NearestNeighbor
    Bilinear
    BicubicCardinal
    BicubicBSpline
    BicubicBCSpline
    Bell
    Gaussian
    WindowedSinc
End Enum

Enum eWindows
    wBartlett
    wBlackman
    wBlackmanHarris
    wBohman
    wConnes
    wCosine
    wGaussian
    wHamming
    wHann
    wKaiser
    wLanczos
    wParzen
    wRectangular
    wWelch
End Enum

Const pi = 3.14159265358979

Public cubic_a#, cubic_b#, cubic_c#, param_d#
Public kernel_size&, sinc_window&, stair_level&

' Simple Box function
Function Box_func(ByVal X#) As Double
    If Abs(X) < 0.5 Then Box_func = 1
End Function

' Linear function with various extent
Function Bartlett_Linear(ByVal X#) As Double
    X = Abs(X)
    If X < kernel_size Then Bartlett_Linear = (1 - Abs(X) / kernel_size) / kernel_size
End Function

' Cardinal cubic spline function
' Depends on one variable: a
' These splines are derived from BC-splines with B=0 and C=-a.
' Only for a=-0.5 the Taylor series expansion of the interpolating function is optimal.
' This cardinal spline is then called Catmull-Rom spline.
' If a=0 then the spline is called Hermite.
Function Cardinal_Cubic_Spline(ByVal X#) As Double
    X = Abs(X)
    If X < 1 Then
        Cardinal_Cubic_Spline = (cubic_a + 2) * X * X * X - (cubic_a + 3) * X * X + 1
    ElseIf X < 2 Then
        Cardinal_Cubic_Spline = cubic_a * X * X * X - 5 * cubic_a * X * X + 8 * cubic_a * X - 4 * cubic_a
    End If
End Function

' Cubic B-spline function
' Depends on no variable.
' It is derived from BC-splines with B=1 and C=0.
Function Cubic_BSpline(ByVal X#) As Double
    Dim a#, B#, c#, d#, tmp#
    If X < 2 Then
        tmp = X + 2
        If tmp > 0 Then a = tmp * tmp * tmp
        tmp = X + 1
        If tmp > 0 Then B = 4 * tmp * tmp * tmp
        If X > 0 Then c = 6 * X * X * X
        tmp = X - 1
        If tmp > 0 Then d = 4 * tmp * tmp * tmp
        Cubic_BSpline = (a - B + c - d) / 6
    End If
End Function

' Cubic BC-spline function
' Mitchell and Netravali derived a family of such cubic filters dependent on two variables: B, C
' Some of the values for B and C correspond to well-known filters,
' e.g., B=1 and C=0 corresponds to the cubic B-spline,
' and C=0 results in the family of Duff's tensioned B-splines.
' Setting B=0 and C=-a results in the family of the cardinal splines which were derived by Keys in 1981.
' Using Taylor series expansion they determined that, numerically, the filters for which B + 2 * C = 1 with 0 <= B <= 1
' are the most accurate within that class
' and that the reconstruction error for synthetic examples is proportional to the square of the sampling distance.
' Two new filters were proposed, the first with B=3/2 and C=1/3 suppresses post-aliasing but is unnecessarily blurring,
' the second with B=1/3 and C=1/3 turns out to be a satisfactory compromise between ringing, blurring, and anisotropy.
Function Cubic_BCSpline(ByVal X#) As Double
    X = Abs(X)
    If X < 1 Then
        Cubic_BCSpline = ((12 - 9 * cubic_b - 6 * cubic_c) * X * X * X + (-18 + 12 * cubic_b + 6 * cubic_c) * X * X + 6 - 2 * cubic_b) / 6
    ElseIf X < 2 Then
        Cubic_BCSpline = ((-cubic_b - 6 * cubic_c) * X * X * X + (6 * cubic_b + 30 * cubic_c) * X * X + (-12 * cubic_b - 48 * cubic_c) * X + (8 * cubic_b + 24 * cubic_c)) / 6
    End If
End Function

' Bell function
Function Bell_Func(ByVal X#) As Double
    X = Abs(X)
    If X < 0.5 Then
        Bell_Func = 0.75 - X * X
    ElseIf X < 1.5 Then
        Bell_Func = 0.5 * (X - 1.5) * (X - 1.5)
    End If
End Function

' Gaussian function
' Could generate very blurry output.
' The wider the function is (higher standard deviation),
' more blurry the image is. This is useful for removing
' noise and aliasing but it won't preserve details.
Function Gaussian_Func(ByVal X#) As Double
    ' 0.398942280401433 = 1 / Sqr(2 * pi)
    Dim o#
    If Abs(X) < kernel_size Then
        o = kernel_size / pi ' standard deviation - could be changed
        Gaussian_Func = 0.398942280401433 / o * Exp(-X * X / (o * o * 2))
    End If
End Function

' Windowed Sinc function
' There's a whole class of reconstruction filters created by starting out
' with the assumption that the sinc function:
'    Sinc(X) = Sin(pi * X) / (pi * X)
' is the perfect image reconstruction filter. It provides the best
' retention of the frequencies you want, and the best attenuation of the
' frequencies that you don't want (because they would cause aliasing).
' Unfortunately, it's impossible to use directly because it is infinite in extent.
' So, practical filters are created by taking the sinc function and
' multiplying it by a "window" function that gradually tapers the Sinc
' function to zero, giving an overall filter with finite size.
Function Windowed_Sinc(ByVal X#) As Double
    If X = 0 Then
        Windowed_Sinc = 1
    Else
        Windowed_Sinc = Window_Func(X) * Sin(pi * X) / (pi * X)
    End If
End Function

' Finite Impulse Response filters (FIR)
Function Window_Func(ByVal X#) As Double
    Dim I#
    If X = 0 Then
        Window_Func = 1
    ElseIf Abs(X) < kernel_size Then
        Select Case sinc_window
            Case wBartlett
                Window_Func = 1 - Abs(X) / kernel_size
            Case wBlackman
                I = pi * X / kernel_size
                Window_Func = 0.42 + 0.5 * Cos(I) + 0.08 * Cos(2 * I)
            Case wBlackmanHarris
                I = pi * X / kernel_size
                Window_Func = 0.42323 + 0.49755 * Cos(I) + 0.07922 * Cos(2 * I)
            Case wBohman
                I = X / kernel_size
                Window_Func = (1 - Abs(I)) * Cos(pi * I) + 1 / pi * Sin(pi * Abs(I))
            Case wConnes
                I = 1 - X * X / (kernel_size * kernel_size)
                Window_Func = I * I
            Case wCosine
                Window_Func = Cos(X * pi / (kernel_size * 2))
            Case wGaussian
                I = kernel_size * 0.4
                Window_Func = 2 ^ (-X * X / (I * I))
            Case wHamming
                Window_Func = 0.54 + 0.46 * Cos(pi * X / kernel_size)
            Case wHann
                Window_Func = 0.5 + 0.5 * Cos(pi * X / kernel_size)
            Case wKaiser
                Window_Func = BesselI0(param_d * Sqr(1 - X * X / (kernel_size * kernel_size))) / BesselI0(param_d)
            Case wLanczos
                I = pi * X / kernel_size
                Window_Func = Sin(I) / I
            Case wParzen
                I = Abs(X) * 1.8 / kernel_size
                If I < 1 Then
                    Window_Func = (4 - 6 * I * I + 3 * I * I * I) * 0.25
                ElseIf I < 2 Then
                    I = 2 - I
                    Window_Func = I * I * I * 0.25
                End If
            Case wRectangular
                Window_Func = 1
            Case wWelch
                Window_Func = 1 - X * X / (kernel_size * kernel_size)
        End Select
    End If
End Function

' Zero Order Modified Bessel Function of the First Kind
Function BesselI0(ByVal X#) As Double
    Dim I#
    ' The expansion usable up to x=20
    I = X * X ' X ^ 2
    BesselI0 = 1 + 0.25 * I
    I = I * I ' X ^ 4
    BesselI0 = BesselI0 + 0.015625 * I
    I = I * X * X ' X ^ 6
    BesselI0 = BesselI0 + 4.34027777777778E-04 * I
    I = I * X * X ' X ^ 8
    BesselI0 = BesselI0 + 6.78168402777778E-06 * I
    I = I * X * X ' X ^ 10
    BesselI0 = BesselI0 + 6.78168402777778E-08 * I
    I = I * X * X ' X ^ 12
    BesselI0 = BesselI0 + 4.7095027970679E-10 * I
    I = I * X * X ' X ^ 14
    BesselI0 = BesselI0 + 2.40280754952444E-12 * I
    I = I * X * X ' X ^ 16
    BesselI0 = BesselI0 + 9.38596699032984E-15 * I
    I = I * X * X ' X ^ 18
    BesselI0 = BesselI0 + 2.89690339207711E-17 * I
    I = I * X * X ' X ^ 20
    BesselI0 = BesselI0 + 7.24225848019278E-20 * I
    I = I * X * X ' X ^ 22
    BesselI0 = BesselI0 + 1.49633439673405E-22 * I
    I = I * X * X ' X ^ 24
    BesselI0 = BesselI0 + 2.59780277210772E-25 * I
    I = I * X * X ' X ^ 26
    BesselI0 = BesselI0 + 3.84290350903508E-28 * I
    I = I * X * X ' X ^ 28
    BesselI0 = BesselI0 + 4.90166263907536E-31 * I
End Function

Sub Resample_NearestNeighbor(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, kX#, kY#
    ' Compute scales on X and Y axes
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    ' Go through destination lines
    For Y = dstHeight - 1 To 0 Step -1
        Y1 = Y * kY ' Nearest value (rounded to integer position)
        ' Go through destination points (pixels)
        For X = 0 To dstWidth - 1
            X1 = X * kX ' Nearest value
            X1 = X1 * 3 ' 24-bit = 3 bytes
            ' Read nearest point from source and set destination point
            bDibD(X * 3, Y) = bDibS(X1, Y1) ' Blue
            bDibD(X * 3 + 1, Y) = bDibS(X1 + 1, Y1) ' Green
            bDibD(X * 3 + 2, Y) = bDibS(X1 + 2, Y1) ' Red
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_Bilinear(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY ' Exact position (floating-point number)
        Y1 = Int(fY) ' Integer position (integer part of number)
        fY = fY - Y1 ' Fraction part of number (integer+fraction=exact)
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            ' Uses various kernel size
            For M = -kernel_size + 1 To kernel_size
                r1 = Bartlett_Linear(M - fY)
                For N = -kernel_size + 1 To kernel_size
                    r2 = Bartlett_Linear(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_BicubicCardinal(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            ' Computes 1 point from 16 (4x4) surrounding points.
            For M = -1 To 2
                ' Parameter for cubic functions is distance of surrounding point and fraction part of number.
                r1 = Cardinal_Cubic_Spline(M - fY)
                For N = -1 To 2
                    r2 = Cardinal_Cubic_Spline(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            If iB < 0 Then iB = 0 ' Check possible bad RGB values
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_BicubicBSpline(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            For M = -1 To 2
                r1 = Cubic_BSpline(M - fY)
                For N = -1 To 2
                    r2 = Cubic_BSpline(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            bDibD(X * 3, Y) = iB     ' We have no need to check values for this filter
            bDibD(X * 3 + 1, Y) = iG ' because it is blurry (averaging).
            bDibD(X * 3 + 2, Y) = iR ' This is also main disadvantage of B-splines.
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_BicubicBCSpline(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            For M = -1 To 2
                r1 = Cubic_BCSpline(M - fY)
                For N = -1 To 2
                    r2 = Cubic_BCSpline(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            If iB < 0 Then iB = 0
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_Bell(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            For M = -1 To 2
                r1 = Bell_Func(M - fY)
                For N = -1 To 2
                    r2 = Bell_Func(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_Gaussian(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            ' Uses various kernel size
            For M = -kernel_size + 1 To kernel_size
                r1 = Gaussian_Func(M - fY)
                For N = -kernel_size + 1 To kernel_size
                    r2 = Gaussian_Func(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            If iB < 0 Then iB = 0
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

Sub Resample_Sinc(bDibD() As Byte, ByVal dstWidth&, ByVal dstHeight&, bDibS() As Byte, ByVal srcWidth&, ByVal srcHeight&)
    Dim X&, Y&, X1&, Y1&, M&, N&, kX#, kY#, fX#, fY#
    Dim iR#, iG#, iB#, r1#, r2#
    kX = (srcWidth - 1) / (dstWidth - 1)
    kY = (srcHeight - 1) / (dstHeight - 1)
    For Y = dstHeight - 1 To 0 Step -1
        fY = Y * kY
        Y1 = Int(fY)
        fY = fY - Y1
        For X = 0 To dstWidth - 1
            fX = X * kX
            X1 = Int(fX)
            fX = fX - X1
            X1 = X1 * 3
            iR = 0: iG = 0: iB = 0
            ' Uses various kernel size
            For M = -kernel_size + 1 To kernel_size
                r1 = Windowed_Sinc(M - fY)
                For N = -kernel_size + 1 To kernel_size
                    r2 = Windowed_Sinc(fX - N)
                    iB = iB + bDibS(X1 + N * 3, Y1 + M) * r1 * r2
                    iG = iG + bDibS(X1 + N * 3 + 1, Y1 + M) * r1 * r2
                    iR = iR + bDibS(X1 + N * 3 + 2, Y1 + M) * r1 * r2
                Next
            Next
            If iB < 0 Then iB = 0
            If iG < 0 Then iG = 0
            If iR < 0 Then iR = 0
            If iB > 255 Then iB = 255
            If iG > 255 Then iG = 255
            If iR > 255 Then iR = 255
            bDibD(X * 3, Y) = iB
            bDibD(X * 3 + 1, Y) = iG
            bDibD(X * 3 + 2, Y) = iR
        Next
        If frmMain.ShowProgress(dstHeight - Y, dstHeight) = True Then Exit Sub
    Next
End Sub

' Normalization is useful especially for blurry interpolations
' like B-spline, Bell, Gaussian and some BC-spline with non-zero B parameter,
' e.g., filters with Y < 1 for X = 0 in Spatial Domain.
Sub Normalize(bDib() As Byte, ByVal nWidth&, ByVal nHeight&)
    Dim X&, Y&, V As Byte, K&, H&(255)
    Dim Min As Byte, Max As Byte
    ' calculate histogram of interpolated image
    For Y = nHeight - 1 To 0 Step -1
        For X = 0 To (nWidth - 1) * 3
            V = bDib(X, Y)
            H(V) = H(V) + 1
        Next
        If frmMain.ShowProgress(nHeight - Y, nHeight) = True Then Exit Sub
    Next
    ' Value 0.001 controls bounds of the histogram (tolerance).
    ' If it is zero then we have strict normalization
    ' but you often can't see any difference between
    ' original and normalized image.
    K = 0.001 * nWidth * nHeight
    Min = 96 ' the highest minimal value
    Max = 159 ' the lowest maximal value
    ' Get Min and Max values by given tolerance K
    For X = 0 To Min - 1
        If H(X) > K Then Min = X: Exit For
    Next
    For X = 255 To Max + 1 Step -1
        If H(X) > K Then Max = X: Exit For
    Next
    If Min = 0 And Max = 255 Then
        MsgBox "The image doesn't need normalization. [Aborted]", vbInformation
        Exit Sub
    End If
    ' Calculate lookup table with normalized values to accelerate processing
    For X = 0 To 255
        K = (X - Min) / (Max - Min) * 255
        If K < 0 Then K = 0
        If K > 255 Then K = 255
        H(X) = K
    Next
    ' Transform output image by lookup table
    For Y = nHeight - 1 To 0 Step -1
        For X = 0 To (nWidth - 1) * 3
            bDib(X, Y) = H(bDib(X, Y))
        Next
        If frmMain.ShowProgress(nHeight - Y, nHeight) = True Then Exit Sub
    Next
End Sub

Sub DoResample(ByVal nType As eTypes, hBitmapDst As Picture, hBitmapSrc As Picture)
    Dim bDibD() As Byte, bDibS() As Byte, bDibT() As Byte
    Dim tSAD As SAFEARRAY2D, tSAS As SAFEARRAY2D
    Dim tBMD As BITMAP, tBMS As BITMAP
    GetObjectAPI hBitmapDst, Len(tBMD), tBMD
    With tSAD ' Array header structure
        .cbElements = 1
        .cDims = 2
        .Bounds(0).cElements = tBMD.bmHeight
        .Bounds(1).cElements = tBMD.bmWidthBytes ' (Width*3 aligned to 4)
        .pvData = tBMD.bmBits ' Pointer to array (bitmap)
    End With
    ' Associate header with array (no need of copying large blocks, direct access)
    CopyMemory ByVal VarPtrArray(bDibD), VarPtr(tSAD), 4
    GetObjectAPI hBitmapSrc, Len(tBMS), tBMS
    With tSAS
        .cbElements = 1
        .cDims = 2
        .Bounds(0).cElements = tBMS.bmHeight
        .Bounds(1).cElements = tBMS.bmWidthBytes
        .pvData = tBMS.bmBits
    End With
    CopyMemory ByVal VarPtrArray(bDibS), VarPtr(tSAS), 4
    Select Case nType
        Case NearestNeighbor
            Resample_NearestNeighbor bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibS, tBMS.bmWidth, tBMS.bmHeight
        Case Bilinear
            ' Extend image by given size
            ReDim bDibT(-3 * (kernel_size - 1) To (tBMS.bmWidth - 1 + kernel_size) * 3 + 2, -(kernel_size - 1) To tBMS.bmHeight - 1 + kernel_size)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2 ' Copy source to extended with edge filling
            Resample_Bilinear bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case BicubicCardinal
            ' Now we need 1 more row of points at top and left and 2 rows at right and bottom otherwise program will crash
            ReDim bDibT(-3 To tBMS.bmWidth * 3 + 5, -1 To tBMS.bmHeight + 1)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_BicubicCardinal bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case BicubicBSpline
            ReDim bDibT(-3 To tBMS.bmWidth * 3 + 5, -1 To tBMS.bmHeight + 1)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_BicubicBSpline bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case BicubicBCSpline
            ReDim bDibT(-3 To tBMS.bmWidth * 3 + 5, -1 To tBMS.bmHeight + 1)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_BicubicBCSpline bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case Bell
            ReDim bDibT(-3 To tBMS.bmWidth * 3 + 5, -1 To tBMS.bmHeight + 1)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_Bell bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case Gaussian
            ReDim bDibT(-3 * (kernel_size - 1) To (tBMS.bmWidth - 1 + kernel_size) * 3 + 2, -(kernel_size - 1) To tBMS.bmHeight - 1 + kernel_size)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_Gaussian bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
        Case WindowedSinc
            ReDim bDibT(-3 * (kernel_size - 1) To (tBMS.bmWidth - 1 + kernel_size) * 3 + 2, -(kernel_size - 1) To tBMS.bmHeight - 1 + kernel_size)
            CopyImage24 bDibS, bDibT, (tBMS.bmWidth - 1) * 3 + 2
            Resample_Sinc bDibD, tBMD.bmWidth, tBMD.bmHeight, bDibT, tBMS.bmWidth, tBMS.bmHeight
    End Select
    CopyMemory ByVal VarPtrArray(bDibD), 0&, 4 ' Important under WinNT platform
    CopyMemory ByVal VarPtrArray(bDibS), 0&, 4 ' Free arrays
End Sub
