Attribute VB_Name = "HiveModule"
Option Explicit

' Error Messages
Public Enum HiveErrors
    [Invalid Index] = &H1
    [Key not Found] = &H2
    [Key Cannot Be Integer] = &H4
    [Duplicate Key] = &H8
    [Invalid Parameter] = &H100
    [Key Cannot be Blank or Zero] = &H1000
End Enum

