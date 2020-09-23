VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "frmPreview.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWindow 
      Caption         =   "Only &Window"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Paint()
    Dim X#, Y#, pY#, Xmax&, Ymax&, I&, J&, DI#
    Dim Size&, nStep#, S$, tRect As RECT
    chkWindow.Visible = False
    Cls
    Xmax = ScaleWidth
    Ymax = ScaleHeight
    J = 0.75 * Ymax
    Line (0, J)-(Xmax, J), vbBlue ' X axis
    Line (CLng(Xmax * 0.5), 0)-(CLng(Xmax * 0.5), Ymax), vbBlue ' Y axis
    Line (0, J)-(0, J) ' Start drawing point
    Select Case frmMain.cboResampler.ListIndex
        Case NearestNeighbor
            Size = 2: nStep = Box_func(0)
            For X = 0 To Xmax
                Y = Box_func((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax ' definite integral
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case Bilinear
            Size = kernel_size + 1: nStep = Bartlett_Linear(0)
            For X = 0 To Xmax
                Y = Bartlett_Linear((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
            lblStatus.Caption = "Definite Integral = 1 <Optimal>"
        Case BicubicCardinal
            Size = 3: nStep = Cardinal_Cubic_Spline(0)
            For X = 0 To Xmax
                Y = Cardinal_Cubic_Spline((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case BicubicBSpline
            Size = 3: nStep = Cubic_BSpline(0)
            For X = 0 To Xmax
                Y = Cubic_BSpline((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case BicubicBCSpline
            Size = 3: nStep = Cubic_BCSpline(0)
            For X = 0 To Xmax
                Y = Cubic_BCSpline((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case Bell
            Size = 3: nStep = Bell_Func(0)
            For X = 0 To Xmax
                Y = Bell_Func((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case Gaussian
            Size = kernel_size + 1: nStep = Gaussian_Func(0)
            For X = 0 To Xmax
                Y = Gaussian_Func((X / Xmax - 0.5) * Size * 2)
                DI = DI + (Y + pY) * Size / Xmax
                pY = Y
                Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                Line -(CLng(X), CLng(Y)), vbRed
            Next
        Case WindowedSinc
            Size = kernel_size + 1: nStep = Windowed_Sinc(0)
            If chkWindow.Value = 1 Then
                For X = 0 To Xmax ' Draw only window function
                    Y = Window_Func((X / Xmax - 0.5) * Size * 2)
                    Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                    Line -(CLng(X), CLng(Y)), vbRed
                Next
            Else
                For X = 0 To Xmax ' Draw Sinc multiplied with window function
                    Y = Windowed_Sinc((X / Xmax - 0.5) * Size * 2)
                    DI = DI + (Y + pY) * Size / Xmax
                    pY = Y
                    Y = ((1 - Y) * 0.5 + 0.25) * Ymax
                    Line -(CLng(X), CLng(Y)), vbRed
                Next
            End If
            chkWindow.Visible = True
    End Select
    Caption = "Spatial domain - Y = " & Round(nStep, 5) & " for X = 0"
    ' Draw marks and numbers on X axis
    tRect.Top = J + 3
    tRect.Bottom = Ymax
    nStep = (Size * 2) * 30 / Xmax
    Y = 0.125
    Do Until nStep < Y
        Y = Round(Y * 2, 2)
    Loop
    nStep = Y
    For X = 0 To Size - 0.1 Step nStep
        X = Round(X, 3)
        I = Xmax / (Size * 2) * (X + Size)
        If X = 0 Then
            I = I + 7
        Else
            Line (I, J - 2)-(I, J + 3), vbBlue ' positive numerical mark
            Line (Xmax - I, J - 2)-(Xmax - I, J + 3), vbBlue ' negative numerical mark
            S = -X
            tRect.Left = Xmax - I - 20
            tRect.Right = Xmax - I + 22
            DrawText hDC, S, Len(S), tRect, 1 ' Draw negative number
        End If
        S = X
        tRect.Left = I - 20
        tRect.Right = I + 22
        DrawText hDC, S, Len(S), tRect, 1 ' Draw positive number
    Next
    ' Draw marks and numbers on Y axis
    I = Xmax * 0.5
    tRect.Left = I + 6
    tRect.Right = Xmax
    nStep = 60 / Ymax
    X = 0.125
    Do Until nStep < X
        X = Round(X * 2, 2)
    Loop
    nStep = X
    For Y = nStep To 1.4 Step nStep
        J = Ymax * (0.75 - 0.5 * Y)
        S = Y
        tRect.Top = J - 20
        tRect.Bottom = J + 20
        DrawText hDC, S, Len(S), tRect, &H24 ' positive number
        Line (I - 2, J)-(I + 3, J), vbBlue ' Mark
        If Y < 0.4 Then
            J = 1.5 * Ymax - J
            S = -Y
            tRect.Top = J - 20
            tRect.Bottom = J + 20
            DrawText hDC, S, Len(S), tRect, &H24 ' negative number
            Line (I - 2, J)-(I + 3, J), vbBlue ' Mark
        End If
    Next
    If frmMain.cboResampler.ListIndex = WindowedSinc And chkWindow.Value = 1 Then
        lblStatus.Caption = ""
    Else
        If Abs(DI - 1) <= 0.005 Then
            S = " <Optimal>"
        Else
            S = " <Unoptimal>"
            If DI > 1 Then S = S & "<Light Output>" Else S = S & "<Dark Output>"
        End If
        lblStatus.Caption = "Definite Integral = " & Round(DI, 3) & S
    End If
    Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = True: Beep
End Sub

Private Sub Form_Resize()
    Paint
End Sub

Private Sub chkWindow_Click()
    Paint
End Sub
