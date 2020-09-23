VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErrorCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type ErrorInformation
    sDescription As String
    sHelpContext  As String
    sHelpFile As String
    lNumber As Long
    sSource As String
End Type

Private Errors() As ErrorInformation

Public lErrors As Long

Public Sub Clear()
    Erase Errors
    lErrors = 0
End Sub

Public Function IsError() As Boolean
    If Err <> 0 Then
        Add Err.Number, Err.Source, Err.Description, Err.HelpContext, Err.HelpFile
        IsError = True
    Else
        IsError = False
    End If
End Function

Public Sub Add( _
    Optional lErrNumber As Long = 0, _
    Optional sSource As String, _
    Optional sDescription As String, _
    Optional sHelpContext As String, _
    Optional sHelpFile As String _
    )

    If lErrNumber = 0 Then
        If Err.Number = 0 Then Exit Sub
        lErrNumber = Err.Number
    End If
    ReDim Preserve Errors(lErrors)
    With Errors(lErrors)
        .sDescription = sDescription
        .sHelpContext = sHelpContext
        .sHelpFile = sHelpFile
        .lNumber = lErrNumber
        .sSource = sSource
    End With
    lErrors = lErrors + 1
End Sub


Public Sub Remove(Optional Index As Variant)
    Dim i As Long
    
    If IsMissing(Index) Then
        If lErrors > 0 Then
            ReDim Preserve Errors(lErrors - 1)
            lErrors = lErrors - 1
        End If
    ElseIf IsNumeric(Index) Then
        If Index < 0 Or Index > lErrors - 1 Then Exit Sub
        For i = Index To lErrors - 1
            Errors(i) = Errors(i + 1)
        Next
        ReDim Preserve Errors(lErrors - 1)
        lErrors = lErrors - 1
    End If
    
End Sub

Public Property Get ErrNumber(Optional Index As Variant) As Long
    If IsMissing(Index) Then
        If lErrors > 0 Then Index = lErrors - 1
    ElseIf IsNumeric(Index) Then
        ErrNumber = Errors(Index).lNumber
    End If
End Property

Public Property Get Description(Optional Index As Variant) As Long
    If IsMissing(Index) Then
        If lErrors > 0 Then Index = lErrors - 1
    ElseIf IsNumeric(Index) Then
        Description = Errors(Index).sDescription
    End If
End Property

Public Property Get HelpContext(Optional Index As Variant) As Long
    If IsMissing(Index) Then
        If lErrors > 0 Then Index = lErrors - 1
    ElseIf IsNumeric(Index) Then
        HelpContext = Errors(Index).sHelpContext
    End If
End Property

Public Property Get HelpFile(Optional Index As Variant) As Long
    If IsMissing(Index) Then
        If lErrors > 0 Then Index = lErrors - 1
    ElseIf IsNumeric(Index) Then
        HelpFile = Errors(Index).sHelpFile
    End If
End Property

Public Property Get Source(Optional Index As Variant) As Long
    If IsMissing(Index) Then
        If lErrors > 0 Then Index = lErrors - 1
    ElseIf IsNumeric(Index) Then
        Source = Errors(Index).sSource
    End If
End Property

Private Sub Class_Terminate()
    Clear
End Sub

Public Property Get Count() As Long
    Count = lErrors
End Property



