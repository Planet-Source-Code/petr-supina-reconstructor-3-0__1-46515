VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3735
   ClientLeft      =   690
   ClientTop       =   975
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   Tag             =   "Reconstructor 3.0"
   Begin VB.CommandButton cmdNormalize 
      Caption         =   "&Normalize"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame fraDestination 
      Caption         =   "Destination Image"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
      Begin VB.CheckBox chkAspect 
         Caption         =   "Constrain &Proportions"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save 24bpp Destination Image"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Text            =   "256"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Text            =   "256"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Height:"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdResample 
      Caption         =   "&Resample"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Frame fraResampling 
      Caption         =   "Interpolation"
      Height          =   3495
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   3375
      Begin VB.HScrollBar hscS 
         Height          =   220
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   21
         Top             =   3120
         Value           =   1
         Width           =   2175
      End
      Begin VB.HScrollBar hsc2 
         Height          =   220
         LargeChange     =   5
         Left            =   120
         Max             =   100
         TabIndex        =   15
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cboSub 
         Height          =   315
         ItemData        =   "frmMain.frx":1042
         Left            =   120
         List            =   "frmMain.frx":1070
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.HScrollBar hsc1 
         Height          =   220
         Left            =   120
         Max             =   20
         Min             =   1
         TabIndex        =   13
         Top             =   600
         Value           =   1
         Width           =   2175
      End
      Begin VB.PictureBox picBC 
         Height          =   2160
         Left            =   120
         Picture         =   "frmMain.frx":1159
         ScaleHeight     =   140
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   140
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "BC-spline scheme (drag to set the control point)"
         Top             =   600
         Width           =   2160
      End
      Begin VB.ComboBox cboResampler 
         Height          =   315
         ItemData        =   "frmMain.frx":F74B
         Left            =   120
         List            =   "frmMain.frx":F767
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblS 
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblSc 
         Caption         =   "Stair Interpolation:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lbl3 
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         ToolTipText     =   "One and only parameter for Cardinal spline. Corresponds to -C."
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lbl2 
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl1 
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source Image"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdSource 
         Caption         =   ">>"
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cboSource 
         Height          =   315
         ItemData        =   "frmMain.frx":F829
         Left            =   120
         List            =   "frmMain.frx":F82B
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "<Choose One>"
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblSource 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Reconstructor by Peter Scale 2003

' This sample application shows some kinds of area interpolation.
' It is not optimized for speed but for understanding.
' Due to no effective optimizations (discrete filters with lookup table)
' these resamplers are reference (continuous), e.g., very strict.

Option Explicit

' structure for OpenFile/SaveFile dialog
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustomFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As String
    lpstrFileTitle As String
    nMaxFileTitle As String
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' API functions to show OpenFile dialog
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
' get state of a key
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Integer

Dim bChange As Boolean

Function ShowProgress(ByVal nPos&, ByVal nMax&) As Boolean
    frmDestination.Refresh
    frmDestination.Line (0, nPos)-(frmDestination.ScaleWidth, nPos), vbRed
    Caption = FormatPercent(nPos / nMax, 0) & " " & Tag & " [hold <Escape> to abort]"
    ShowProgress = GetAsyncKeyState(vbKeyEscape) And &HF000
End Function

Private Sub cboResampler_Click()
    ' Enable or disable controls we need
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
    picBC.Visible = False
    hsc1.Visible = False
    hsc2.Visible = False
    cboSub.Visible = False
    Select Case cboResampler.ListIndex
        Case Bilinear
            hsc1.Value = 1
            hsc1_Scroll
            hsc1.Visible = True
            lbl1.Visible = True
        Case BicubicCardinal
            ' This will set a=-0.5 (C=0.5; B=0)
            picBC_MouseMove vbLeftButton, 0, 73, 121
            lbl1.Visible = True
            lbl2.Visible = True
            lbl3.Visible = True
            picBC.Visible = True
        Case BicubicBSpline
            ' B=1; C=0
            picBC_MouseMove vbLeftButton, 0, 18, 11
            lbl1.Visible = True
            lbl2.Visible = True
            picBC.Visible = True
        Case BicubicBCSpline
            ' B=1/3; C=1/3
            picBC_MouseMove vbLeftButton, 0, 55, 85
            lbl1.Visible = True
            lbl2.Visible = True
            picBC.Visible = True
        Case Gaussian
            hsc1.Value = 2
            hsc1_Scroll
            hsc1.Visible = True
            lbl1.Visible = True
        Case WindowedSinc
            cboSub.ListIndex = wBlackmanHarris
            hsc1.Value = 3
            hsc1_Scroll
            cboSub.Visible = True
            hsc1.Visible = True
            lbl1.Visible = True
    End Select
    frmPreview.Paint
End Sub

Private Sub cboSource_Click()
    Dim tBM As BITMAP, sPic As StdPicture
    Dim CDC&, CDC1&
    On Error GoTo Out
    With frmSource
        ' Load chosen picture
        .Picture = LoadPicture(App.Path & "\Images\" & cboSource.Text)
        ' Get informations about loaded bitmap
        GetObjectAPI .Picture, Len(tBM), tBM
        ' Show info about source
        lblSource.Caption = "Width: " & tBM.bmWidth & "    Height: " & tBM.bmHeight & "    BPP: " & tBM.bmBitsPixel
        .Caption = "Source Image: " & tBM.bmWidth & "x" & tBM.bmHeight
        ' If non 24bpp image loaded convert to it
        If tBM.bmBitsPixel <> 24 Then
            ' Create 24bpp empty (black) image
            Set sPic = CreatePicture(tBM.bmWidth, tBM.bmHeight, 24)
            CDC = CreateCompatibleDC(0) ' Temporary devices
            CDC1 = CreateCompatibleDC(0)
            DeleteObject SelectObject(CDC, .Picture) ' Source bitmap
            DeleteObject SelectObject(CDC1, sPic) ' Converted bitmap
            ' Copy between two different depths
            BitBlt CDC1, 0, 0, tBM.bmWidth, tBM.bmHeight, CDC, 0, 0, vbSrcCopy
            DeleteDC CDC: DeleteDC CDC1 ' Erase devices
            .Picture = sPic ' Set visible image
        End If
        txtWidth_Change
        .Move Left, Top + Height, .Width - (.ScaleWidth - tBM.bmWidth) * Screen.TwipsPerPixelX, .Height - (.ScaleHeight - tBM.bmHeight) * Screen.TwipsPerPixelY
        frmDestination.Move .Left + .Width, .Top
        .Show vbModeless, Me
        SetFocus
        cmdResample.Enabled = True
    End With
Out:
End Sub

Private Sub cboSub_Click()
    sinc_window = cboSub.ListIndex
    If sinc_window = wKaiser Then
        hsc2.Value = 22
        hsc2.Visible = True
        lbl2.Visible = True
    Else
        hsc2.Visible = False
        lbl2.Visible = False
    End If
    frmPreview.Paint
End Sub

Private Sub cmdNormalize_Click()
    Dim bDibD() As Byte
    Dim tSAD As SAFEARRAY2D, tBMD As BITMAP
    GetObjectAPI frmDestination.Picture, Len(tBMD), tBMD
    With tSAD ' Array header structure
        .cbElements = 1
        .cDims = 2
        .Bounds(0).cElements = tBMD.bmHeight
        .Bounds(1).cElements = tBMD.bmWidthBytes ' (Width*3 aligned to 4)
        .pvData = tBMD.bmBits ' Pointer to array (bitmap)
    End With
    ' Associate header with array (no need of copying large blocks, direct access)
    CopyMemory ByVal VarPtrArray(bDibD), VarPtr(tSAD), 4
    Normalize bDibD, tBMD.bmWidth, tBMD.bmHeight
    CopyMemory ByVal VarPtrArray(bDibD), 0&, 4
    frmDestination.Refresh
    Caption = Tag
End Sub

Private Sub cmdResample_Click()
    Dim tBM As BITMAP, sPic As StdPicture, I&, dw&, dH&, W&, H&, pW&, pH&
    ' Check if destination dimensions are correct
    If IsNumeric(txtWidth) = False Or IsNumeric(txtHeight) = False Then GoTo Out
    If txtWidth < 2 Or txtWidth > 2048 Or txtHeight < 2 Or txtHeight > 2048 Then GoTo Out
    frmDestination.Move frmSource.Left + frmSource.Width, frmSource.Top
    frmDestination.Show vbModeless, Me
    SetFocus
    ' Resample using Stair Interpolation
    ' This is helping for large size changes
    ' otherwise use Level 1 (don't use with Nearest Neighbor).
    ' It could remove aliasing but also could cause other artifacts!
    GetObjectAPI frmSource.Picture, Len(tBM), tBM
    dw = (txtWidth - tBM.bmWidth) / stair_level
    dH = (txtHeight - tBM.bmHeight) / stair_level
    W = tBM.bmWidth: H = tBM.bmHeight
    GetObjectAPI frmDestination.Picture, Len(tBM), tBM
    For I = 1 To stair_level
        If I = stair_level Then
            W = txtWidth
            H = txtHeight
        Else
            W = W + dw
            H = H + dH
        End If
        If W <= 0 Then W = 1
        If H <= 0 Then H = 1
        If pW <> W Or pH <> H Or I = 1 Then
        pW = W: pH = H
        If I = 1 Then
            Set sPic = frmSource.Picture
        Else
            Set sPic = frmDestination.Picture
        End If
        If Not (I = 1 And tBM.bmWidth = W And tBM.bmHeight = H) Then
            frmDestination.Picture = CreatePicture(W, H, 24)
        End If
        frmDestination.Move frmSource.Left + frmSource.Width, frmSource.Top, frmDestination.Width - (frmDestination.ScaleWidth - W) * Screen.TwipsPerPixelX, frmDestination.Height - (frmDestination.ScaleHeight - H) * Screen.TwipsPerPixelY
        frmDestination.Caption = "Destination Image: " & W & "x" & H
        DoResample cboResampler.ListIndex, frmDestination.Picture, sPic
        End If
        If GetAsyncKeyState(vbKeyEscape) And &HF000 Then Exit For
    Next
    frmDestination.Refresh
    cmdSave.Enabled = True
    cmdNormalize.Enabled = True
    Caption = Tag
    Exit Sub
Out: MsgBox "Invalid destination dimensions", vbCritical
End Sub

Private Sub cmdSave_Click()
    Dim OFN As OPENFILENAME
    With OFN
        .lpstrFile = String$(260, 0)
        .nMaxFile = Len(.lpstrFile)
        .lpstrFilter = "Windows Bitmaps (*.bmp)" & vbNullChar & "*.bmp" & vbNullChar & vbNullChar
        .lpstrDefExt = "bmp"
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .Flags = 6
        .lStructSize = Len(OFN)
        ' SaveFile dialog
        If GetSaveFileName(OFN) = 0 Then Exit Sub
        ' Save image to the file
        SavePicture frmDestination.Picture, Left$(.lpstrFile, InStr(1, .lpstrFile, vbNullChar) - 1)
    End With
End Sub

Private Sub cmdSource_Click()
    Dim OFN As OPENFILENAME, tBM As BITMAP
    Dim sPic As StdPicture, CDC&, CDC1&, S$
    With OFN
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFilter = "Images" & vbNullChar & "*.bmp;*.jpg;*.gif;*.ico;*.cur;*.rle" & vbNullChar & vbNullChar
        .Flags = &H1804
        .lpstrFile = String$(260, 0)
        .nMaxFile = Len(.lpstrFile)
        .lStructSize = Len(OFN)
        If GetOpenFileName(OFN) = 0 Then Exit Sub
        S = Left$(.lpstrFile, InStr(1, .lpstrFile, vbNullChar) - 1)
    End With
    cboSource.Text = Mid$(S, InStrRev(S, "\") + 1)
    On Error GoTo Out
    With frmSource
        .Picture = LoadPicture(S)
        GetObjectAPI .Picture, Len(tBM), tBM
        lblSource.Caption = "Width: " & tBM.bmWidth & "    Height: " & tBM.bmHeight & "    BPP: " & tBM.bmBitsPixel
        .Caption = "Source Image: " & tBM.bmWidth & "x" & tBM.bmHeight
        If tBM.bmBitsPixel <> 24 Then
            Set sPic = CreatePicture(tBM.bmWidth, tBM.bmHeight, 24)
            CDC = CreateCompatibleDC(0)
            CDC1 = CreateCompatibleDC(0)
            DeleteObject SelectObject(CDC, .Picture)
            DeleteObject SelectObject(CDC1, sPic)
            BitBlt CDC1, 0, 0, tBM.bmWidth, tBM.bmHeight, CDC, 0, 0, vbSrcCopy
            DeleteDC CDC: DeleteDC CDC1
            .Picture = sPic
        End If
        txtWidth_Change
        .Move Left, Top + Height, .Width - (.ScaleWidth - tBM.bmWidth) * Screen.TwipsPerPixelX, .Height - (.ScaleHeight - tBM.bmHeight) * Screen.TwipsPerPixelY
        .Show vbModeless, Me
        SetFocus
        cmdResample.Enabled = True
    End With
Out:
End Sub

Private Sub Form_Load()
    Dim strFile$
    ' Select resampler in the list
    cboResampler.ListIndex = BicubicBCSpline
    ' Load names of available images into ComboBox
    strFile = Dir(App.Path & "\Images\*.*")
    While Len(strFile)
        cboSource.AddItem strFile
        strFile = Dir
    Wend
    Caption = Tag
    hscS_Scroll
    Show
    frmPreview.Move Left + Width, Top
    frmPreview.Show vbModeless, Me
    SetFocus
End Sub

Private Sub hsc1_Change()
    hsc1_Scroll
End Sub

Private Sub hsc1_Scroll()
    kernel_size = hsc1.Value
    lbl1.Caption = "Extent: " & kernel_size
    lbl1.Refresh
    frmPreview.Paint
End Sub

Private Sub hsc2_Change()
    hsc2_Scroll
End Sub

Private Sub hsc2_Scroll()
    param_d = hsc2.Value * 0.25
    lbl2.Caption = "par = " & Round(param_d, 1)
    lbl2.Refresh
    frmPreview.Paint
End Sub

Private Sub hscS_Change()
    hscS_Scroll
End Sub

Private Sub hscS_Scroll()
    stair_level = hscS.Value
    lblS.Caption = "Level = " & stair_level
End Sub

Private Sub picBC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBC_MouseMove Button, Shift, X, Y
End Sub

Private Sub picBC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        picBC.Cls
        ' Determine max and min values by picture and resampler
        If X < 18 Then X = 18
        If Y < 11 Then Y = 11
        If X > 128 Then X = 128
        If Y > 121 Then Y = 121
        Select Case cboResampler.ListIndex
            Case BicubicBSpline
                X = 18: Y = 11
            Case BicubicCardinal
                Y = 121
        End Select
        ' Draw circle around the point
        picBC.Circle (X, Y), 3, &HFF00FF
        ' Set a, B, C and show them
        cubic_c = (X - 18) / 110
        cubic_a = -cubic_c
        lbl3.Caption = "a = " & Round(cubic_a, 2)
        lbl2.Caption = "C = " & Round(cubic_c, 2)
        cubic_b = 1 - (Y - 11) / 110
        lbl1.Caption = "B = " & Round(cubic_b, 2)
        Refresh
        frmPreview.Paint
    End If
End Sub

Private Sub picBC_Paint()
    picBC_MouseMove vbLeftButton, 0, cubic_c * 113 + 16, (1 - cubic_b) * 113 + 9
End Sub

Private Sub txtHeight_Change()
    Dim tBM As BITMAP
    If chkAspect.Value = 1 And IsNumeric(txtHeight) = True And frmSource.Picture <> 0 And bChange = False Then
        If txtHeight >= 2 And txtHeight <= 2048 Then
            GetObjectAPI frmSource.Picture, Len(tBM), tBM
            bChange = True
            txtWidth = CLng(txtHeight / tBM.bmHeight * tBM.bmWidth)
            bChange = False
        End If
    End If
End Sub

Private Sub txtWidth_Change()
    Dim tBM As BITMAP
    If chkAspect.Value = 1 And IsNumeric(txtWidth) = True And frmSource.Picture <> 0 And bChange = False Then
        If txtWidth >= 2 And txtWidth <= 2048 Then
            GetObjectAPI frmSource.Picture, Len(tBM), tBM
            bChange = True
            txtHeight = CLng(txtWidth / tBM.bmWidth * tBM.bmHeight)
            bChange = False
        End If
    End If
End Sub
