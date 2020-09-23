VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const csClassName As String = "CHive"

' Default behaviours of Hive
Private Const DefaultInitialAlloc As Long = 100
Private Const DefaultGrowthFactor As Double = 1.5

Private GrowthFactor As Double
Private InitialAlloc As Long

' Sentinel is Node(0)
Private Const Sentinel As Long = 0

' Node Color
Private Enum EColor
    Black
    Red
End Enum

' fields associated with each node
Private Type ItemData
  lLeft As Long          ' Left child
  lRight As Long         ' Right child
  lParent As Long        ' Parent
  Color As EColor        ' red or black
  vKey As Variant        ' item key
  vData As Variant       ' item data
End Type

Private Items() As ItemData     ' Array which stored all item

' support for FindFirst and FindNext
Private StackIndex As Integer
Private Stack(1 To 32) As Long
Private NextNode As Long

Private Root As Long            ' root of binary tree
Private Node As CNode           ' class for allocating nodes

Private lCount As Long          ' No of items
Private lIndex() As Long        ' Used to map items to index

Private mCompareMode As VbCompareMethod

Public Errors As CErrorCollection   ' Contains all the errors

Public RaiseError As Boolean    ' True: Call Err.Raise; False: Don't
Public AllowDuplicate As Boolean ' True: Allow duplicate; False: Don't

' GUID Key generation code
Private Declare Function CoCreateGuid Lib _
    "OLE32.DLL" (pGuid As GUID) As Long
Private Declare Function StringFromGUID2 Lib _
    "OLE32.DLL" (pGuid As GUID, _
    ByVal PointerToString As Long, _
    ByVal MaxLength As Long) As Long

' GUID Result
Private Const GUID_OK As Long = 0

' Structure to hold GUID
Private Type GUID
    Guid1           As Long             ' 32 bit
    Guid2           As Integer          ' 16 bit
    Guid3           As Integer          ' 16 bit
    Guid4(0 To 7)   As Byte             ' 64 bit
End Type

' For Raw memory copy
Private Declare Sub MoveMemory Lib "kernel32" _
      Alias "RtlMoveMemory" (dest As Long, _
    Source As Long, ByVal numBytes As Long)
    
Public Function CreateGUIDKey() As String
    Const GUID_LENGTH   As Long = 38

    Dim udtGUID As GUID
    'User Defined Type
    Dim strFormattedGUID As String
    'The formatted string
    Dim lngResult As Long
    'Useless result flag

    ' Create a GUID
    lngResult = CoCreateGuid(udtGUID)

    If lngResult = GUID_OK Then
        ' Pre-allocate space for the ID
        strFormattedGUID = String$(GUID_LENGTH, 0)
        ' Convert the GUID to a
        'formatted string
        StringFromGUID2 udtGUID, _
            StrPtr(strFormattedGUID), GUID_LENGTH + 1
    Else
        strFormattedGUID = ""
    End If
    CreateGUIDKey = strFormattedGUID
End Function

Private Sub Raise(ByVal errno As HiveErrors, ByVal sLocation As String)
'   inputs:
'       errno       Error ID
'       sLocation   Location of the error
    
    Dim sErrMsg As String
    
    Select Case errno
        
        Case [Invalid Index]
            sErrMsg = "Invalid Index."
        Case [Key not Found]
            sErrMsg = "Key not Found in the Hive."
        Case [Key Cannot Be Integer]
            sErrMsg = "Key cannot be Integer or Long or Byte or any kind of Fixed Digit."
        Case [Duplicate Key]
            sErrMsg = "Duplicate Key."
        Case [Invalid Parameter]
            sErrMsg = "Invalid Parameter."

    End Select
        
    If RaiseError Then
        Err.Raise vbObjectError + 5000 + errno, csClassName + "." + sLocation, sErrMsg
    End If
    If Errors Is Nothing Then
      Set Errors = New CErrorCollection
    End If
    Errors.Add vbObjectError + 5000 + errno, sLocation, sErrMsg
    
End Sub

Private Function GetKeyIndex(ByVal vKey As Variant) As Long
'   inputs:
'       vKeysVal          vKeys of the node
'   returns:
'       the index of the item in the Array

    Select Case VarType(vKey)
        Case vbByte, vbInteger, vbLong
            If vKey < 0 Or vKey > lCount Then
                Raise [Invalid Index], "GetIndex"
            Else
                GetKeyIndex = lIndex(vKey)
            End If
        Case Else
            GetKeyIndex = FindNode(vKey)
            If GetKeyIndex = 0 Then Raise [Key not Found], "FindNode"
    End Select
End Function

Private Function GetIndex(ByVal KeyIndex As Long) As Long
    Dim i As Long
    
    For i = 1 To lCount
        If lIndex(i) = KeyIndex Then
            GetIndex = i
            Exit Function
        End If
    Next
End Function
Private Function FindNode(ByVal KeyVal As Variant) As Long
'   inputs:
'       Key                   ' designates key to find
'   returns:
'       index to node
'   action:
'       Search tree for designated key, and return index to node.
'   errors:
'       Key Not Found
'
    Dim current As Long

    ' find node specified by key
    current = Root
    
    ' ------------------------------------
    ' if compare mode is binary
    ' then match exact key otherwise
    ' ignore case if key is a string
    ' ------------------------------------
    If mCompareMode <> vbBinaryCompare And VarType(KeyVal) = vbString Then
      KeyVal = LCase(KeyVal)
      Do While current <> Sentinel
          If LCase(Items(current).vKey) = KeyVal Then
              FindNode = current
              Exit Function
          Else
              If KeyVal < LCase(Items(current).vKey) Then
                  current = Items(current).lLeft
              Else
                  current = Items(current).lRight
              End If
          End If
      Loop
    Else
      Do While current <> Sentinel
          If Items(current).vKey = KeyVal Then
              FindNode = current
              Exit Function
          Else
              If KeyVal < Items(current).vKey Then
                  current = Items(current).lLeft
              Else
                  current = Items(current).lRight
              End If
          End If
      Loop
      
    End If
End Function

Private Sub RotateLeft(ByVal x As Long)
'   inputs:
'       x                     designates node
'   action:
'       perform a lLeft tree rotation about "x"
'
    Dim y As Long

    ' rotate node x to lLeft

    y = Items(x).lRight

    ' establish x.lRight link
    Items(x).lRight = Items(y).lLeft
    If Items(y).lLeft <> Sentinel Then Items(Items(y).lLeft).lParent = x

    ' establish y.lParent link
    If y <> Sentinel Then Items(y).lParent = Items(x).lParent
    If Items(x).lParent <> 0 Then
        If x = Items(Items(x).lParent).lLeft Then
            Items(Items(x).lParent).lLeft = y
        Else
            Items(Items(x).lParent).lRight = y
        End If
    Else
        Root = y
    End If

    ' link x and y
    Items(y).lLeft = x
    If x <> Sentinel Then Items(x).lParent = y
End Sub

Private Sub RotateRight(ByVal x As Long)
'   inputs:
'       x                     designates node
'   action:
'       perform a lRight tree rotation about "x"
'
    Dim y As Long

    ' rotate node x to lRight

    y = Items(x).lLeft

    ' establish x.lLeft link
    Items(x).lLeft = Items(y).lRight
    If Items(y).lRight <> Sentinel Then Items(Items(y).lRight).lParent = x

    ' establish y.lParent link
    If y <> Sentinel Then Items(y).lParent = Items(x).lParent
    If Items(x).lParent <> 0 Then
        If x = Items(Items(x).lParent).lRight Then
            Items(Items(x).lParent).lRight = y
        Else
            Items(Items(x).lParent).lLeft = y
        End If
    Else
        Root = y
    End If

    ' link x and y
    Items(y).lRight = x
    If x <> Sentinel Then Items(x).lParent = y
End Sub

Private Sub InsertFixup(ByRef x As Long)
'   inputs:
'       x                     designates node
'   action:
'       maintains red-black tree properties after inserting node x
'
    Dim y As Long

    Do While x <> Root
        If Items(Items(x).lParent).Color <> Red Then Exit Do
        ' we have a violation
        If Items(x).lParent = Items(Items(Items(x).lParent).lParent).lLeft Then
            y = Items(Items(Items(x).lParent).lParent).lRight
            If Items(y).Color = Red Then

                ' uncle is Red
                Items(Items(x).lParent).Color = Black
                Items(y).Color = Black
                Items(Items(Items(x).lParent).lParent).Color = Red
                x = Items(Items(x).lParent).lParent
            Else

                ' uncle is Black
                If x = Items(Items(x).lParent).lRight Then
                    ' make x a lLeft child
                    x = Items(x).lParent
                    RotateLeft x
                End If

                ' recolor and rotate
                Items(Items(x).lParent).Color = Black
                Items(Items(Items(x).lParent).lParent).Color = Red
                RotateRight Items(Items(x).lParent).lParent
            End If
        Else

            ' mirror image of above code
            y = Items(Items(Items(x).lParent).lParent).lLeft
            If Items(y).Color = Red Then

                ' uncle is Red
                Items(Items(x).lParent).Color = Black
                Items(y).Color = Black
                Items(Items(Items(x).lParent).lParent).Color = Red
                x = Items(Items(x).lParent).lParent
            Else

                ' uncle is Black
                If x = Items(Items(x).lParent).lLeft Then
                    x = Items(x).lParent
                    RotateRight x
                End If
                Items(Items(x).lParent).Color = Black
                Items(Items(Items(x).lParent).lParent).Color = Red
                RotateLeft Items(Items(x).lParent).lParent
            End If
        End If
    Loop
    Items(Root).Color = Black
End Sub

Public Function Add(ByRef Item As Variant, Optional ByVal Key As Variant, Optional ByVal Before As Variant, Optional ByVal After As Variant)
Attribute Add.VB_MemberFlags = "40"
'   inputs:
'       Item        Item to store
'       Key         Key to use
'       Before      The item before which this item will be inserted
'       After      The item After which this item will be inserted
'   action:
'       Inserts Item with Key.
'   error:
'       [Duplicate Key]
'
    Dim current As Long
    Dim p As Long
    Dim x As Long
    Dim i As Long
    Dim j As Long
    Dim lItems As Long
    Dim strTempKey As String  ' Used to store lcase key
          
    ' Validate Key
    If IsMissing(Key) Then
        Key = CreateGUIDKey()
    Else
        Select Case VarType(Key)
            Case vbLong, vbInteger, vbByte
                Raise [Key Cannot Be Integer], "Add"
                Exit Function
                
            Case vbString
                If Key = "" Then
                    Raise [Key Cannot be Blank or Zero], "Add"
                    Exit Function
                End If
                
        End Select
    End If
    ' allocate node for data and insert in tree
    If Node Is Nothing Then Init InitialAlloc, GrowthFactor
    
    ' find where node belongs
    current = Root
    p = 0
    
    ' ---------------------------------------------------------------
    ' Search hive if the key already exist. If exist then if duplicate
    ' allowed then accept otherwise get out. After serching look for a
    ' position where the new items key will be stored in the Red-Black
    ' tree. Thank you.
    ' ---------------------------------------------------------------
    If VarType(Key) = vbString Then strTempKey = LCase(Key)
    Do While current <> Sentinel
        If mCompareMode <> vbBinaryCompare And VarType(Key) = vbString Then
          
          If LCase(Items(current).vKey) = strTempKey Then
              If Not AllowDuplicate Then
                  Raise [Duplicate Key], "Add"
                  Exit Function
              End If
          End If
          
          p = current
          If strTempKey < LCase(Items(current).vKey) Then
              current = Items(current).lLeft
          Else
              current = Items(current).lRight
          End If
          
        Else
          If Items(current).vKey = Key Then
            If Not AllowDuplicate Then
                Raise [Duplicate Key], "Add"
                Exit Function
            End If
          End If
          
          p = current
          If Key < Items(current).vKey Then
              current = Items(current).lLeft
          Else
              current = Items(current).lRight
          End If
        End If
        
        
    Loop

    ' setup new node
    x = Node.Alloc()
    lItems = UBound(Items())
    If x > lItems Then
        ReDim Preserve Items(0 To lItems * GrowthFactor)
        ReDim Preserve lIndex(1 To (lItems * GrowthFactor) + 2)
    End If
    
    Items(x).lParent = p
    Items(x).lLeft = Sentinel
    Items(x).lRight = Sentinel
    Items(x).Color = Red
    
    ' Increase the counter. Increased value is
    ' required below
    lCount = lCount + 1
    ' Adjust position
    If Not IsMissing(Before) Then
        Before = GetKeyIndex(Before)
        If Before = 0 Then
            Raise [Key not Found], "Add"
            Exit Function
        End If
        i = GetIndex(Before)
        InsertItem i
        lIndex(i) = x
        
    ElseIf Not IsMissing(After) Then
        After = GetKeyIndex(After)
        If After = 0 Then
            Raise [Key not Found], "Add"
            Exit Function
        End If
        i = GetIndex(After) + 1
        InsertItem i
        lIndex(i) = x
    Else
        lIndex(lCount) = x
    End If
    
    ' copy fields to node
    Items(x).vKey = Key
    If IsObject(Item) Then
        Set Items(x).vData = Item
    Else
        Items(x).vData = Item
    End If

    ' insert node in tree
    If p <> 0 Then
        If mCompareMode <> vbBinaryCompare And VarType(Key) = vbString Then
          If strTempKey < LCase(Items(p).vKey) Then
              Items(p).lLeft = x
          Else
              Items(p).lRight = x
          End If
        Else
          If Key < Items(p).vKey Then
            Items(p).lLeft = x
          Else
            Items(p).lRight = x
          End If
        End If
    Else
        Root = x
    End If

    InsertFixup x
End Function

Private Sub DeleteFixup(ByRef x As Long)
'   inputs:
'       x                     designates node
'   action:
'       maintains red-black tree properties after deleting a node
'
    Dim w As Long

    Do While (x <> Root)
        If Items(x).Color <> Black Then Exit Do
        If x = Items(Items(x).lParent).lLeft Then
            w = Items(Items(x).lParent).lRight
            If Items(w).Color = Red Then
                Items(w).Color = Black
                Items(Items(x).lParent).Color = Red
                RotateLeft Items(x).lParent
                w = Items(Items(x).lParent).lRight
            End If
            If Items(Items(w).lLeft).Color = Black _
            And Items(Items(w).lRight).Color = Black Then
                Items(w).Color = Red
                x = Items(x).lParent
            Else
                If Items(Items(w).lRight).Color = Black Then
                    Items(Items(w).lLeft).Color = Black
                    Items(w).Color = Red
                    RotateRight w
                    w = Items(Items(x).lParent).lRight
                End If
                Items(w).Color = Items(Items(x).lParent).Color
                Items(Items(x).lParent).Color = Black
                Items(Items(w).lRight).Color = Black
                RotateLeft Items(x).lParent
                x = Root
            End If
        Else
            w = Items(Items(x).lParent).lLeft
            If Items(w).Color = Red Then
                Items(w).Color = Black
                Items(Items(x).lParent).Color = Red
                RotateRight Items(x).lParent
                w = Items(Items(x).lParent).lLeft
            End If
            If Items(Items(w).lRight).Color = Black _
            And Items(Items(w).lLeft).Color = Black Then
                Items(w).Color = Red
                x = Items(x).lParent
            Else
                If Items(Items(w).lLeft).Color = Black Then
                    Items(Items(w).lRight).Color = Black
                    Items(w).Color = Red
                    RotateLeft w
                    w = Items(Items(x).lParent).lLeft
                End If
                Items(w).Color = Items(Items(x).lParent).Color
                Items(Items(x).lParent).Color = Black
                Items(Items(w).lLeft).Color = Black
                RotateRight Items(x).lParent
                x = Root
            End If
        End If
    Loop
    Items(x).Color = Black
End Sub

Public Function Remove(ByVal KeyVal As Variant) As Long
'   inputs:
'       KeyVal                key of node to delete
'   action:
'       Deletes record with key KeyVal.
'   error:
'       errKeyNotFound
'
    Dim x As Long
    Dim y As Long
    Dim z As Long
    Dim i As Long
    
    z = GetKeyIndex(KeyVal) ' FindNode(KeyVal)
    If z = 0 Then
        Raise [Invalid Index], "Remove"
    End If
    
    '  delete node z from tree
    If Items(z).lLeft = Sentinel Or Items(z).lRight = Sentinel Then
        ' y has a Sentinel node as a child
        y = z
    Else
        ' find tree successor with a Sentinel node as a child
        y = Items(z).lRight
        Do While Items(y).lLeft <> Sentinel
            y = Items(y).lLeft
        Loop
    End If

    ' x is y's only child, and x may be a sentinel node
    If Items(y).lLeft <> Sentinel Then
        x = Items(y).lLeft
    Else
        x = Items(y).lRight
    End If

    ' remove y from the lParent chain
    Items(x).lParent = Items(y).lParent
    If Items(y).lParent <> 0 Then
        If y = Items(Items(y).lParent).lLeft Then
            Items(Items(y).lParent).lLeft = x
        Else
            Items(Items(y).lParent).lRight = x
        End If
    Else
        Root = x
    End If
    If y <> z Then
        Dim j As Long
        ' copy data fields from y to z
        ' z item now contains y item
        Items(z).vKey = Items(y).vKey
        If IsObject(Items(y).vData) Then
          Set Items(z).vData = Items(y).vData
        Else
          Items(z).vData = Items(y).vData
        End If
        
        ' Swap index of z and y
        i = GetIndex(z)
        j = GetIndex(y)
        
        lIndex(i) = y
        lIndex(j) = z
        
        
    End If
    
    ' if we removed a black node, we need to do some fixup
    If Items(y).Color = Black Then DeleteFixup x

    Set Items(y).vData = Nothing
    Items(y).vData = Empty
    Set Items(y).vKey = Nothing
    Items(y).vKey = Empty
    
    ' Delete index of y
    i = GetIndex(y)
    LiftItem i
    Remove = i
    
    lIndex(lCount) = 0
    lCount = lCount - 1
    
    Node.Free y
End Function

Private Function GetNextNode() As Long
'   returns:
'       index to next node, 0 if none
'   action:
'       Finds index to next node.
'
    Do While (NextNode <> 0 Or StackIndex <> 0)
        Do While NextNode <> 0
            StackIndex = StackIndex + 1
            Stack(StackIndex) = NextNode
            NextNode = Items(NextNode).lLeft
        Loop
        GetNextNode = Stack(StackIndex)
        StackIndex = StackIndex - 1
        NextNode = Items(GetNextNode).lRight
        Exit Function
    Loop
    Raise [Key not Found], "GetNextNode"
End Function


Public Function FindFirst(ByRef KeyVal As Variant) As Variant
'   outputs:
'       KeyVal                key of node to find
'   returns:
'       record associated with key
'   action:
'       For sequential access, finds first record.
'   errors:
'       errKeyNotFound
'
    Dim n As Long
    
    ' for sequential access, call FindFirst, followed by
    ' repeated calls to FindNext
    
    NextNode = Root
    n = GetNextNode()
    KeyVal = Items(n).vKey
    
    If IsObject(Items(n).vData) Then
        Set FindFirst = Items(n).vData
    Else
        FindFirst = Items(n).vData
    End If
End Function

Public Function FindNext(ByRef KeyVal As Variant) As Variant
'   outputs:
'       KeyVal                record key
'   returns:
'       record associated with key
'   action:
'       For sequential access, finds next record.
'   errors:
'       errKeyNotFound
'
    Dim n As Long
    
    ' for sequential access, call FindFirst, followed by
    ' repeated calls to FindNext
    
    n = GetNextNode()
    KeyVal = Items(n).vKey
    
    
    If IsObject(Items(n).vData) Then
        Set FindNext = Items(n).vData
    Else
        FindNext = Items(n).vData
    End If
    
End Function

Public Sub Init( _
        ByVal InitialAllocVal As Long, _
        ByVal GrowthFactorVal As Single)
'   inputs:
'       InitialAllocVal         initial value for allocating nodes
'       GrowthFactorVal         amount to grow node storage space
'   action:
'       initialize tree
'
    GrowthFactor = GrowthFactorVal

    ' allocate nodes
    ReDim Items(0 To InitialAllocVal)
    ReDim lIndex(1 To InitialAllocVal + 1)
    
    ' initialize root and sentinel
    Items(Sentinel).lLeft = Sentinel
    Items(Sentinel).lRight = Sentinel
    Items(Sentinel).lParent = 0
    Items(Sentinel).Color = Black
    Root = Sentinel

    ' startup node manager
    Set Node = New CNode
    Node.Init InitialAllocVal, GrowthFactorVal
    
    ' Initialize error container
    Set Errors = New CErrorCollection
    
    StackIndex = 0
    lCount = 0
End Sub

Public Function Clear() As Long
'   action:
'       Clears memory
    Dim i As Long
    
    If Node Is Nothing Then Exit Function
    Set Node = Nothing
    Set Errors = Nothing
    
    For i = 1 To lCount - 1
        Set Items(i).vData = Nothing
        Items(i).vData = Empty
    Next
    
    For i = 1 To lCount - 1
        Set Items(i).vKey = Nothing
        Items(i).vKey = Empty
    Next
    
    Erase Items
    lCount = 0
    Clear = lCount
End Function

Public Function Exist(ByVal vKey As Variant) As Boolean
'   action:
'       Searches in the array for the specified item
'   inputs:
'       vKey        The key or Index of the item
'   returns:
'       True is item exist. Otherwise false
    
    Exist = GetKeyIndex(vKey) > 0
    
End Function

Private Sub Class_Initialize()
    InitialAlloc = DefaultInitialAlloc
    GrowthFactor = DefaultGrowthFactor
    mCompareMode = vbTextCompare
    
End Sub

Private Sub Class_Terminate()
'   action:
'       release memory
'
    Clear
    Set Errors = Nothing
End Sub

Public Property Get Item(ByVal vKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
'   action:
'       Returns the item specified in vKey
'   inputs:
'       vKey        The key or Index of the item

    Dim lIndex As Long
    
    Set Item = Nothing
    lIndex = GetKeyIndex(vKey)
    If lIndex > 0 Then
        If IsObject(Items(lIndex).vData) Then
            Set Item = Items(lIndex).vData
        Else
            Item = Items(lIndex).vData
        End If
    Else
      Raise [Key not Found], "Item"
    End If
    
End Property

Public Property Let Item(ByVal vKey As Variant, ByRef vNewValue As Variant)
    Dim lIndex As Long
    
    lIndex = GetKeyIndex(vKey)
    If lIndex > 0 Then
        If IsObject(vNewValue) Then
            Set Items(lIndex).vData = vNewValue
        Else
            Items(lIndex).vData = vNewValue
        End If
    End If
End Property

Public Property Get Key(ByVal Index As Long) As Variant
    If Index < 0 Or Index > lCount Then
        Raise [Invalid Index], "Key[Read]"
    Else
        Key = Items(lIndex(Index)).vKey
    End If
End Property

Public Property Let Key(ByVal Index As Long, ByVal vNewKey As Variant)
    If Index < 0 Or Index > lCount Then
        Raise [Invalid Index], "Key[Assign]"
    Else
        If FindNode(vNewKey) <> 0 Then
            If Not AllowDuplicate Then
                Raise [Duplicate Key], "Key[Assign]"
            Else
                Items(lIndex(Index)).vKey = vNewKey
            End If
        End If
    End If
End Property

Public Property Get Count() As Long
    Count = lCount
End Property

Private Sub LiftItem(i As Long)
    'Dim x As Long
    'For x = i To lCount - 1
    '    lIndex(x) = lIndex(x + 1)
    'Next
    MoveMemory lIndex(i), lIndex(i + 1), (lCount - i) * 4
End Sub

Private Sub InsertItem(i As Long)
  'Dim j As Long
  'For j = lCount To i + 1 Step -1
  '  lIndex(j) = lIndex(j - 1)
  'Next
  MoveMemory lIndex(i + 1), lIndex(i), (lCount - i) * 4
End Sub

Public Property Get CompareMode() As VbCompareMethod
  CompareMode = mCompareMode
End Property

Public Property Let CompareMode(ByVal NewMode As VbCompareMethod)
  mCompareMode = NewMode
End Property
