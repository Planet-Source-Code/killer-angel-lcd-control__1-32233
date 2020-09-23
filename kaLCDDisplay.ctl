VERSION 5.00
Begin VB.UserControl kaLCDDisplay 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   PaletteMode     =   4  'None
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   70
   ToolboxBitmap   =   "kaLCDDisplay.ctx":0000
   Begin VB.Timer tmrProcessCommands 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer tmrWrite 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   240
   End
   Begin VB.Image imgLetters 
      Height          =   135
      Left            =   0
      Picture         =   "kaLCDDisplay.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   8565
   End
End
Attribute VB_Name = "kaLCDDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Default Property Values
Const m_def_AutoSize = 0
Const m_def_Mode = 0
Const m_def_Scrollable = 1
Const m_def_TotalCols = 16
Const m_def_TotalRows = 2
Const m_def_WriteSpeed = 0

'Property Variables
Private m_AutoSize As Boolean
Private m_Mode As LCDModes
Private m_Picture As Picture
Private m_Scrollable As Boolean
Private m_TotalCols As Integer
Private m_TotalRows As Integer
Private m_WriteSpeed As Integer

'Constants
Const letterWidth = 6
Const letterHeight = 9
Const lastChar = 126
Private Separator_Command As String
Private Separator_Write As String

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Private Enum ProcessCommands
    procmdWriteText
    procmdWriteSpeed
    procmdClear
    procmdScroll
End Enum
Public Enum Alignment
    LCDAlignLeft
    LCDAlignCenter
    LCDAlignRight
End Enum
Public Enum LCDModes
    AlphaNumeric
    Graphic
End Enum

'Global Timed Write Variables
Private sTextToWrite As String
Private iTextStart As Integer
Private iTextStop As Integer
Private iRowToWrite As Integer

'Global Variables
Private bypass As Boolean
Private imgDC As Long
Private imgBitmap As Long
Private PicInfo As BITMAP
Private WriteArray() As String

'Class Objects
Private objHive As CHive

'API
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents when in Graphic Mode."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
'    If m_Mode = AlphaNumeric Then Exit Property
    m_AutoSize = New_AutoSize
    DrawControl
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,1,0,0
Public Property Get Mode() As LCDModes
Attribute Mode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal New_Mode As LCDModes)
'    If New_Mode = AlphaNumeric Then
'        m_AutoSize = False
'        Set m_Picture = Nothing
'    End If
    m_Mode = New_Mode
    DrawControl
    PropertyChanged "Mode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
'    If m_Mode = AlphaNumeric Then Exit Property
    Set m_Picture = New_Picture
    DrawControl
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Scrollable() As Boolean
Attribute Scrollable.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Scrollable = m_Scrollable
End Property

Public Property Let Scrollable(ByVal New_Scrollable As Boolean)
    m_Scrollable = New_Scrollable
    PropertyChanged "Scrollable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get TotalCols() As Integer
Attribute TotalCols.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If Ambient.UserMode Then Err.Raise 393
    TotalCols = m_TotalCols
End Property

Public Property Let TotalCols(ByVal New_Cols As Integer)
    If Ambient.UserMode Then Err.Raise 382
    If New_Cols < 1 Then Exit Property
    m_TotalCols = New_Cols
    DrawControl
    PropertyChanged "TotalCols"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,1,0,0
Public Property Get TotalRows() As Integer
Attribute TotalRows.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If Ambient.UserMode Then Err.Raise 393
    TotalRows = m_TotalRows
End Property

Public Property Let TotalRows(ByVal New_Rows As Integer)
    If Ambient.UserMode Then Err.Raise 382
    If New_Rows < 1 Then Exit Property
    m_TotalRows = New_Rows
    DrawControl
    PropertyChanged "TotalRows"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get WriteSpeed() As Integer
Attribute WriteSpeed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    WriteSpeed = m_WriteSpeed
End Property

Public Property Let WriteSpeed(ByVal New_WriteSpeed As Integer)
    Dim bTemp As Boolean
    
    If Not bypass Then
        If objHive.Count = 0 And Not tmrWrite.Enabled Then
            bTemp = True
        Else
            objHive.Add procmdWriteSpeed & Separator_Command & New_WriteSpeed
            bTemp = False
        End If
    Else
        bTemp = True
    End If
    If bTemp Then
        m_WriteSpeed = New_WriteSpeed
        tmrWrite.Interval = IIf(m_WriteSpeed = 0, 1, m_WriteSpeed)
        PropertyChanged "WriteSpeed"
    End If
    bypass = False
End Property

Private Function DeleteMemDC(MemDC As Long, hBitmap As Long) As Long
    DeleteMemDC = DeleteDC(MemDC)
    DeleteObject hBitmap
End Function

Private Function DrawControl()
    On Error Resume Next
    Dim X As Integer, Y As Integer
    Dim mProgress As Long
    
    Select Case m_Mode
        Case AlphaNumeric
            UserControl_Resize
            For Y = 0 To m_TotalRows - 1
                For X = 0 To m_TotalCols - 1
                    BitBlt hDC, letterWidth * X, letterHeight * Y, letterWidth + 1, letterHeight, imgDC, 0, 0, vbSrcCopy
                Next
            Next
            ReDim WriteArray(m_TotalRows - 1): For X = 0 To m_TotalRows - 1: WriteArray(X) = Space(m_TotalCols): Next
        Case Graphic
            Set UserControl.Picture = m_Picture
            UserControl_Resize
            If Not (m_Picture Is Nothing) Then
                BurkeBW Image, 15, mProgress
            End If
    End Select
    Refresh
    tmrWrite.Interval = IIf(m_WriteSpeed = 0, 1, m_WriteSpeed)
End Function

Private Function GenerateMemImgDC(ByRef MemDC As Long, ByRef hBitmap As Long, Letters As Image) As Long
    On Error GoTo Error:
    MemDC = CreateCompatibleDC(0)
    If MemDC = 0 Then
        GenerateMemImgDC = 0
        Exit Function
    End If
    hBitmap = Letters.Picture.Handle
    If hBitmap = 0 Then
        DeleteDC MemDC
        Exit Function
    End If
    SelectObject MemDC, hBitmap
    GenerateMemImgDC = 1
    Exit Function
Error:
    DeleteDC MemDC
    GenerateMemImgDC = 0
End Function

Private Function ScrollLines()
    Dim X As Integer, Y As Integer
    
    If Not tmrWrite.Enabled Then
        If m_TotalRows < 2 Then Exit Function
        For X = 0 To m_TotalRows - 2
            WriteArray(X) = WriteArray(X + 1)
        Next
        WriteArray(m_TotalRows - 1) = Space(m_TotalCols)
        For Y = 0 To m_TotalRows - 1
            For X = 0 To m_TotalCols - 1
                BitBlt hDC, letterWidth * X, letterHeight * Y, letterWidth + 1, letterHeight, imgDC, letterWidth * (Asc(Mid(WriteArray(Y), X + 1, 1)) - 32), 0, vbSrcCopy
            Next
        Next
        Refresh
        iRowToWrite = iRowToWrite - 1
    Else
        objHive.Add procmdScroll & Separator_Command
    End If
End Function

Private Sub TextFixAlign(ByRef sOldText As String, sNewText As String, ByVal Align As Integer, ByVal iCol As Integer)
    Dim sTemp As String, i As Integer
    Dim iSpcSize As Integer
    
    For i = 1 To Len(sNewText)
        sTemp = Mid(sNewText, i, 1)
        If (Asc(sTemp) > lastChar) Then
            sNewText = Replace(sNewText, sTemp, " ")
        End If
    Next
    i = m_TotalCols - Len(sNewText)
    iSpcSize = IIf(i < 0, 0, i)
    sNewText = Left(sNewText, m_TotalCols - iCol)
    Select Case Align
        Case LCDAlignLeft
            sNewText = Left(sOldText, iCol - 1) & sNewText
            If (iSpcSize - iCol + 1) > 0 Then sOldText = sNewText & Right(sOldText, iSpcSize - iCol + 1)
        Case LCDAlignCenter
            i = iSpcSize - (iSpcSize \ 2)
            sOldText = Left(sOldText, iSpcSize \ 2) & sNewText & Right(sOldText, i)
        Case LCDAlignRight
            sOldText = Left(sOldText, iSpcSize - iCol + 1) & sNewText & Right(sOldText, iCol - 1)
    End Select
End Sub

Public Function Clear()
    Dim X As Integer, Y As Integer
    
    If m_Mode = Graphic Then Exit Function
    If Not tmrWrite.Enabled Then
        For X = 0 To m_TotalRows - 1
            WriteArray(X) = Space(m_TotalCols)
        Next
        For Y = 0 To m_TotalRows - 1
            For X = 0 To m_TotalCols - 1
                BitBlt hDC, letterWidth * X, letterHeight * Y, letterWidth + 1, letterHeight, imgDC, 0, 0, vbSrcCopy
            Next
        Next
        Refresh
        iRowToWrite = 0
    Else
        objHive.Add procmdClear & Separator_Command
    End If
End Function

Public Function WriteText(ByVal sText As String, Optional ByVal Align As Integer = 0, Optional ByVal iRow As Integer = 0, Optional ByVal iCol As Integer = 1)
    Dim X As Integer, iTemp As Integer
    
    If m_Mode = Graphic Then Exit Function
    iTemp = iRowToWrite + 1
    If (iTemp < 1) Or (Not m_Scrollable And iTemp > m_TotalRows) Then Exit Function
    If (m_Scrollable) And (iTemp > m_TotalRows) And (iRow <= 0) Then
        ScrollLines
        iTemp = iRowToWrite + 1
    End If
    If m_WriteSpeed = 0 Then
        iRowToWrite = iTemp
        If iRow > 0 Then iRowToWrite = iRow
        TextFixAlign WriteArray(iRowToWrite - 1), sText, Align, iCol
        sText = WriteArray(iRowToWrite - 1)
        For X = 0 To m_TotalCols - 1
            BitBlt hDC, letterWidth * X, letterHeight * (iRowToWrite - 1), letterWidth, letterHeight, imgDC, letterWidth * (Asc(Mid(sText, X + 1, 1)) - 32), 0, vbSrcCopy
        Next
        Refresh
        If objHive.Count > 0 Then tmrProcessCommands.Enabled = True
    Else
        If Not tmrWrite.Enabled Then
            iRowToWrite = iTemp
            If iRow > 0 Then iRowToWrite = iRow
            TextFixAlign WriteArray(iRowToWrite - 1), sText, Align, iCol
            sTextToWrite = sText
            iTextStart = InStr(WriteArray(iRowToWrite - 1), sText)
            iTextStop = iTextStart + Len(sText)
            tmrWrite.Enabled = True
         Else
            objHive.Add procmdWriteText & Separator_Command & sText & Separator_Write & Align & Separator_Write & iRow & Separator_Write & iCol
        End If
    End If
End Function

Private Sub UserControl_Initialize()
    Separator_Command = Chr(255)
    Separator_Write = Chr(254)
    bypass = False
    iRowToWrite = 0
    Set objHive = New CHive
    GenerateMemImgDC imgDC, imgBitmap, imgLetters
End Sub

Private Sub UserControl_InitProperties()
    m_AutoSize = m_def_AutoSize
    m_Mode = m_def_Mode
    Set m_Picture = LoadPicture("")
    m_Scrollable = m_def_Scrollable
    m_TotalCols = m_def_TotalCols
    m_TotalRows = m_def_TotalRows
    m_WriteSpeed = m_def_WriteSpeed
    DrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_Mode = PropBag.ReadProperty("Mode", m_def_Mode)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Scrollable = PropBag.ReadProperty("Scrollable", m_def_Scrollable)
    m_TotalCols = PropBag.ReadProperty("TotalCols", m_def_TotalCols)
    m_TotalRows = PropBag.ReadProperty("TotalRows", m_def_TotalRows)
    m_WriteSpeed = PropBag.ReadProperty("WriteSpeed", m_def_WriteSpeed)
    DrawControl
End Sub

Private Sub UserControl_Resize()
    Static bResizing As Boolean
    
    If bResizing Then Exit Sub
    bResizing = True
    Width = (m_TotalCols * letterWidth + 1) * Screen.TwipsPerPixelX
    Height = m_TotalRows * letterHeight * Screen.TwipsPerPixelY
    Select Case m_Mode
        Case AlphaNumeric
        Case Graphic
            If m_AutoSize And Not (m_Picture Is Nothing) Then
                Call GetObject(m_Picture, Len(PicInfo), PicInfo)
                Width = PicInfo.bmWidth * Screen.TwipsPerPixelX
                Height = PicInfo.bmHeight * Screen.TwipsPerPixelY
            End If
    End Select
    bResizing = False
    Refresh
End Sub

Private Sub UserControl_Terminate()
    DeleteMemDC imgDC, imgBitmap
    Set objHive = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Mode", m_Mode, m_def_Mode)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Scrollable", m_Scrollable, m_def_Scrollable)
    Call PropBag.WriteProperty("TotalCols", m_TotalCols, m_def_TotalCols)
    Call PropBag.WriteProperty("TotalRows", m_TotalRows, m_def_TotalRows)
    Call PropBag.WriteProperty("WriteSpeed", m_WriteSpeed, m_def_WriteSpeed)
End Sub

Private Sub tmrWrite_Timer()
    Dim iTemp As Integer
    
    iTemp = InStr(WriteArray(iRowToWrite - 1), sTextToWrite) - 1
    If iTextStart >= iTextStop Then
        tmrWrite.Enabled = False
        tmrProcessCommands.Enabled = True
        Exit Sub
    End If
    BitBlt hDC, letterWidth * (iTextStart - 1), letterHeight * (iRowToWrite - 1), letterWidth, letterHeight, imgDC, letterWidth * (Asc(Mid(sTextToWrite, iTextStart - iTemp, 1)) - 32), 0, vbSrcCopy
    iTextStart = iTextStart + 1
    Refresh
End Sub

Private Sub tmrProcessCommands_Timer()
    Dim iCommand As Integer, sTemp As String
    Dim aTemp() As String
    
    If objHive.Count = 0 Then
        tmrProcessCommands.Enabled = False
        Exit Sub
    End If
    aTemp = Split(objHive.Item(1), Separator_Command)
    iCommand = Val(aTemp(0))
    sTemp = aTemp(1)
    objHive.Remove 1
    Select Case iCommand
        Case procmdWriteText
            tmrProcessCommands.Enabled = False
            aTemp = Split(sTemp, Separator_Write)
            WriteText aTemp(0), Val(aTemp(1)), Val(aTemp(2)), Val(aTemp(3))
        Case procmdWriteSpeed
            bypass = True
            WriteSpeed = Val(sTemp)
        Case procmdClear
            Clear
        Case procmdScroll
            ScrollLines
    End Select
End Sub
