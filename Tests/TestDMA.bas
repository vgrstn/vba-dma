Attribute VB_Name = "TestDMA"
'@IgnoreModule UseMeaningfulName, VariableNotAssigned, UnassignedVariableUsage, LineLabelNotUsed, VariableNotUsed, AssignmentNotUsed, IntegerDataType, FunctionReturnValueDiscarded
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Rubberduck unit tests for DMA: every result is checked against plain VBA or the RtlMoveMemory API.")
Option Compare Database
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Rubberduck unit tests for DMA
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' Every public member of DMA has at least one test. The expected results come
' from plain VBA (assignments, VarPtr, LBound/UBound, Collection methods) or
' from the RtlMoveMemory API, never from DMA itself. The Poke tests also check
' that the bytes next to the target are left alone. The timings are in
' BenchmarkDMA, not here.

#Const LateBinding = False

#If LateBinding Then
Private Assert As Object
#Else
Private Assert As AssertClass
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows/win32/devnotes/rtlmovememory
Private Declare PtrSafe Sub RefCopy Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal targetPtr As LongPtr, _
    ByVal sourcePtr As LongPtr, _
    ByVal byteCount As Long _
)

' The VarPtr function of the VBA runtime, declared for arrays.
Private Declare PtrSafe Function RefVarPtrArray Lib "VBE7" Alias "VarPtr" ( _
    ByRef anyArray() As Any _
) As LongPtr

#If Win64 Then
Private Const PtrBytes As Long = 8
Private Const MaxPtr As LongPtr = &H7FFFFFFFFFFFFFFF^
Private Const MinPtr As LongPtr = &H8000000000000000^
Private Const AddressHexDigits As Long = 16
#Else
Private Const PtrBytes As Long = 4
Private Const MaxPtr As LongPtr = &H7FFFFFFF
Private Const MinPtr As LongPtr = &H80000000
Private Const AddressHexDigits As Long = 8
#End If

' The offset of the data in a Variant.
Private Const VariantDataOffset As Long = 8

' The bytes next to a poked value are set to this pattern.
Private Const Filler As Byte = &HA5

Private Const ErrInvalidCall As Long = 5
Private Const ErrOutOfRange As Long = 9
Private Const ErrTypeMismatch As Long = 13
Private Const ErrObjectRequired As Long = 424
Private Const ErrArgNotOptional As Long = 449


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test module set-up
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@ModuleInitialize
Private Sub ModuleInitialize()
#If LateBinding Then
    Set Assert = CreateObject("Rubberduck.AssertClass")
#Else
    Set Assert = New AssertClass
#End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the LongPtr stored at a memory address, read with the API.")
Private Function RefPeekPtr(ByVal memoryPtr As LongPtr) As LongPtr

    Dim Result As LongPtr
    RefCopy VarPtr(Result), memoryPtr, PtrBytes
    RefPeekPtr = Result

End Function


'@Description("Returns a byte buffer filled with the filler pattern.")
Private Function FilledBuffer(ByVal byteCount As Long) As Byte()

    Dim Buffer() As Byte
    ReDim Buffer(0 To byteCount - 1)
    Dim Position As Long
    For Position = 0 To byteCount - 1
        Buffer(Position) = Filler
    Next
    FilledBuffer = Buffer

End Function


'@Description("Returns the number of filler bytes in a buffer outside the range first..first+size-1.")
Private Function FillerOutside( _
    ByRef byteBuffer() As Byte, _
    ByVal firstByte As Long, _
    ByVal byteCount As Long _
) As Long

    Dim Position As Long
    For Position = LBound(byteBuffer) To UBound(byteBuffer)
        If Position < firstByte Or Position >= firstByte + byteCount Then
            If byteBuffer(Position) = Filler Then FillerOutside = FillerOutside + 1
        End If
    Next

End Function


'@Description("Returns the elements of a one-dimensional array separated by commas.")
Private Function ListOf(ByVal arrayValues As Variant) As String

    Dim Element As Variant
    For Each Element In arrayValues
        ListOf = ListOf & "," & Element
    Next
    ListOf = Mid$(ListOf, 2)

End Function


'@Description("Returns a collection with five items; the second item has no key.")
Private Function SampleBag() As Collection

    Dim Bag As Collection
    Set Bag = New Collection
    Bag.Add "one", "A"
    Bag.Add 2
    Bag.Add "three", "C"
    Bag.Add 4, "D"
    Bag.Add "five", "E"
    Set SampleBag = Bag

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: pointers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Pointers")
Private Sub OffsetPtr_AddsOffsetAcrossSignBoundary()
    On Error GoTo TestFail

    Assert.AreEqual CLngPtr(1024), OffsetPtr(1000, 24)
    Assert.AreEqual CLngPtr(992), OffsetPtr(1000, -8)
    Assert.AreEqual CLngPtr(-4), OffsetPtr(-8, 4)
    Assert.AreEqual MinPtr, OffsetPtr(MaxPtr, 1)
    Assert.AreEqual MaxPtr, OffsetPtr(MinPtr, -1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: peek
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Peek")
Private Sub PeekBool_ReadsBooleanAtAddress()
    On Error GoTo TestFail

    Dim Slot As Boolean
    Slot = True
    Assert.AreEqual True, PeekBool(VarPtr(Slot))
    Slot = False
    Assert.AreEqual False, PeekBool(VarPtr(Slot))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekByte_ReadsByteAtAddress()
    On Error GoTo TestFail

    Dim Slots(0 To 2) As Byte
    Slots(0) = Filler
    Slots(2) = Filler
    Dim Sample As Variant
    For Each Sample In Array(0, 1, 127, 128, 255)
        Slots(1) = Sample
        Assert.AreEqual CByte(Sample), _
                        PeekByte(VarPtr(Slots(1))), _
                        "PeekByte " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekCur_ReadsCurrencyAtAddress()
    On Error GoTo TestFail

    Dim Slot As Currency
    Dim Sample As Variant
    For Each Sample In Array(0@, 1.2345@, -123456789.1234@, 922337203685477.5807@, -922337203685477.5807@)
        Slot = Sample
        Assert.AreEqual CCur(Sample), _
                        PeekCur(VarPtr(Slot)), _
                        "PeekCur " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekDbl_ReadsDoubleAtAddress()
    On Error GoTo TestFail

    Dim Slot As Double
    Dim Sample As Variant
    For Each Sample In Array(0#, -1.5, 3.14159265358979, 1E+308, -2.2E-308)
        Slot = Sample
        Assert.AreEqual CDbl(Sample), _
                        PeekDbl(VarPtr(Slot)), _
                        "PeekDbl " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekInt_ReadsIntegerAtAddress()
    On Error GoTo TestFail

    Dim Slot As Integer
    Dim Sample As Variant
    For Each Sample In Array(0, 1, -1, 32767, -32768)
        Slot = Sample
        Assert.AreEqual CInt(Sample), _
                        PeekInt(VarPtr(Slot)), _
                        "PeekInt " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekLng_ReadsLongAtAddress()
    On Error GoTo TestFail

    Dim Slot As Long
    Dim Sample As Variant
    For Each Sample In Array(0&, 1&, -1&, &H7FFFFFFF, &H80000000)
        Slot = Sample
        Assert.AreEqual CLng(Sample), _
                        PeekLng(VarPtr(Slot)), _
                        "PeekLng " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


#If Win64 Then
'@TestMethod("Peek")
Private Sub PeekLngLng_ReadsLongLongAtAddress()
    On Error GoTo TestFail

    Dim Slot As LongLong
    Dim Sample As Variant
    For Each Sample In Array(0^, 1^, -1^, &H7FFFFFFFFFFFFFFF^, &H8000000000000000^)
        Slot = Sample
        Assert.AreEqual CLngLng(Sample), _
                        PeekLngLng(VarPtr(Slot)), _
                        "PeekLngLng " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
#End If


'@TestMethod("Peek")
Private Sub PeekPtr_ReadsLongPtrAtAddress()
    On Error GoTo TestFail

    Dim Slot As LongPtr
    Dim Sample As Variant
    For Each Sample In Array(CLngPtr(0), CLngPtr(1), CLngPtr(-1), MaxPtr, MinPtr)
        Slot = Sample
        Assert.AreEqual CLngPtr(Sample), _
                        PeekPtr(VarPtr(Slot)), _
                        "PeekPtr " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekSng_ReadsSingleAtAddress()
    On Error GoTo TestFail

    Dim Slot As Single
    Dim Sample As Variant
    For Each Sample In Array(0!, -1.5!, 3.141593!, 3.4E+38!, -1.5E-30!)
        Slot = Sample
        Assert.AreEqual CSng(Sample), _
                        PeekSng(VarPtr(Slot)), _
                        "PeekSng " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekVar_ReadsVariantAtAddress()
    On Error GoTo TestFail

    Dim Slot As Variant
    Dim Sample As Variant
    For Each Sample In Array(42&, 2.5, "text", True, #1/2/2026#, 12.34@)
        Slot = Sample
        Dim Actual As Variant: Actual = PeekVar(VarPtr(Slot))
        Assert.AreEqual VarType(Sample), _
                        VarType(Actual), _
                        "PeekVar type " & Sample
        Assert.IsTrue Actual = Sample, "PeekVar " & Sample
    Next

    Slot = Empty
    Assert.IsTrue IsEmpty(PeekVar(VarPtr(Slot))), "PeekVar Empty"
    Slot = Null
    Assert.IsTrue IsNull(PeekVar(VarPtr(Slot))), "PeekVar Null"

    Slot = Array(1, 2, 3)
    Actual = PeekVar(VarPtr(Slot))
    Assert.AreEqual "1,2,3", Join(Actual, ",")

    Dim Bag As Collection
    Set Bag = New Collection
    Set Slot = Bag
    Assert.IsTrue PeekVar(VarPtr(Slot)) Is Bag, "PeekVar object"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekStr_ReadsStringAtStrPtr()
    On Error GoTo TestFail

    Dim Text As String
    Dim Sample As Variant
    For Each Sample In Array("a", "text", ChrW$(&H20AC) & " 100", String$(1000, "x"))
        Text = Sample
        Assert.AreEqual CStr(Sample), _
                        PeekStr(StrPtr(Text)), _
                        "PeekStr " & Left$(Sample, 10)
    Next
    Assert.AreEqual vbNullString, PeekStr(vbNullPtr)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Peek")
Private Sub PeekObj_ReturnsObjectAtObjPtr()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = New Collection
    Bag.Add "item"

    Dim Copy As Object
    Set Copy = PeekObj(ObjPtr(Bag))
    Assert.IsTrue Copy Is Bag, "PeekObj returns the same object"

    ' The copy holds its own reference.
    Set Bag = Nothing
    Assert.AreEqual "item", Copy.Item(1)

    Assert.IsTrue PeekObj(vbNullPtr) Is Nothing, "PeekObj(vbNullPtr) is Nothing"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: poke
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Poke")
Private Sub PokeBool_WritesBooleanOnly()
    On Error GoTo TestFail

    Dim Buffer() As Byte
    Dim Slot As Boolean

    Buffer = FilledBuffer(8)
    PokeBool VarPtr(Buffer(2)), True
    RefCopy VarPtr(Slot), VarPtr(Buffer(2)), 2
    Assert.AreEqual True, Slot
    Assert.AreEqual 6&, FillerOutside(Buffer, 2, 2), "PokeBool True neighbours"

    Buffer = FilledBuffer(8)
    PokeBool VarPtr(Buffer(2)), False
    RefCopy VarPtr(Slot), VarPtr(Buffer(2)), 2
    Assert.AreEqual False, Slot
    Assert.AreEqual 6&, FillerOutside(Buffer, 2, 2), "PokeBool False neighbours"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeByte_WritesByteOnly()
    On Error GoTo TestFail

    Dim Sample As Variant
    For Each Sample In Array(0, 1, 128, 255)
        Dim Buffer() As Byte: Buffer = FilledBuffer(8)
        PokeByte VarPtr(Buffer(3)), Sample
        Assert.AreEqual CByte(Sample), Buffer(3), "PokeByte " & Sample
        Assert.AreEqual 7&, FillerOutside(Buffer, 3, 1), "PokeByte neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeCur_WritesCurrencyOnly()
    On Error GoTo TestFail

    Dim Slot As Currency
    Dim Sample As Variant
    For Each Sample In Array(0@, 1.2345@, -123456789.1234@, 922337203685477.5807@)
        Dim Buffer() As Byte: Buffer = FilledBuffer(16)
        PokeCur VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), 8
        Assert.AreEqual CCur(Sample), Slot, "PokeCur " & Sample
        Assert.AreEqual 8&, FillerOutside(Buffer, 4, 8), "PokeCur neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeDbl_WritesDoubleOnly()
    On Error GoTo TestFail

    Dim Slot As Double
    Dim Sample As Variant
    For Each Sample In Array(0#, -1.5, 3.14159265358979, 1E+308)
        Dim Buffer() As Byte: Buffer = FilledBuffer(16)
        PokeDbl VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), 8
        Assert.AreEqual CDbl(Sample), Slot, "PokeDbl " & Sample
        Assert.AreEqual 8&, FillerOutside(Buffer, 4, 8), "PokeDbl neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeInt_WritesIntegerOnly()
    On Error GoTo TestFail

    Dim Slot As Integer
    Dim Sample As Variant
    For Each Sample In Array(0, -1, 32767, -32768)
        Dim Buffer() As Byte: Buffer = FilledBuffer(8)
        PokeInt VarPtr(Buffer(2)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(2)), 2
        Assert.AreEqual CInt(Sample), Slot, "PokeInt " & Sample
        Assert.AreEqual 6&, FillerOutside(Buffer, 2, 2), "PokeInt neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeLng_WritesLongOnly()
    On Error GoTo TestFail

    Dim Slot As Long
    Dim Sample As Variant
    For Each Sample In Array(0&, -1&, &H7FFFFFFF, &H80000000)
        Dim Buffer() As Byte: Buffer = FilledBuffer(12)
        PokeLng VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), 4
        Assert.AreEqual CLng(Sample), Slot, "PokeLng " & Sample
        Assert.AreEqual 8&, FillerOutside(Buffer, 4, 4), "PokeLng neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


#If Win64 Then
'@TestMethod("Poke")
Private Sub PokeLngLng_WritesLongLongOnly()
    On Error GoTo TestFail

    Dim Slot As LongLong
    Dim Sample As Variant
    For Each Sample In Array(0^, -1^, &H7FFFFFFFFFFFFFFF^, &H8000000000000000^)
        Dim Buffer() As Byte: Buffer = FilledBuffer(16)
        PokeLngLng VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), 8
        Assert.AreEqual CLngLng(Sample), Slot, "PokeLngLng " & Sample
        Assert.AreEqual 8&, FillerOutside(Buffer, 4, 8), "PokeLngLng neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
#End If


'@TestMethod("Poke")
Private Sub PokePtr_WritesLongPtrOnly()
    On Error GoTo TestFail

    Dim Slot As LongPtr
    Dim Sample As Variant
    For Each Sample In Array(CLngPtr(0), CLngPtr(-1), MaxPtr, MinPtr)
        Dim Buffer() As Byte: Buffer = FilledBuffer(16)
        PokePtr VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), PtrBytes
        Assert.AreEqual CLngPtr(Sample), Slot, "PokePtr " & Sample
        Assert.AreEqual 16& - PtrBytes, FillerOutside(Buffer, 4, PtrBytes), "PokePtr neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeSng_WritesSingleOnly()
    On Error GoTo TestFail

    Dim Slot As Single
    Dim Sample As Variant
    For Each Sample In Array(0!, -1.5!, 3.141593!, 3.4E+38!)
        Dim Buffer() As Byte: Buffer = FilledBuffer(12)
        PokeSng VarPtr(Buffer(4)), Sample
        RefCopy VarPtr(Slot), VarPtr(Buffer(4)), 4
        Assert.AreEqual CSng(Sample), Slot, "PokeSng " & Sample
        Assert.AreEqual 8&, FillerOutside(Buffer, 4, 4), "PokeSng neighbours " & Sample
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeVar_ReplacesVariant()
    On Error GoTo TestFail

    Dim Slot As Variant
    Dim Sample As Variant
    For Each Sample In Array(42&, 2.5, "text", True, #1/2/2026#, 12.34@)
        Slot = "previous value"
        PokeVar VarPtr(Slot), Sample
        Assert.AreEqual VarType(Sample), VarType(Slot), "PokeVar type " & Sample
        Assert.IsTrue Slot = Sample, "PokeVar " & Sample
    Next

    PokeVar VarPtr(Slot), Null
    Assert.IsTrue IsNull(Slot), "PokeVar Null"

    PokeVar VarPtr(Slot), Array("a", "b")
    Assert.AreEqual "a,b", Join(Slot, ",")

    Dim Bag As Collection
    Set Bag = New Collection
    PokeVar VarPtr(Slot), Bag
    Assert.IsTrue Slot Is Bag, "PokeVar object"
    PokeVar VarPtr(Slot), 7&
    Assert.AreEqual 7&, Slot

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeStr_OverwritesKeepingLength()
    On Error GoTo TestFail

    Dim Text As String

    Text = "....."
    PokeStr StrPtr(Text), "abcde"
    Assert.AreEqual "abcde", Text

    ' LSet semantics: a shorter value is padded with spaces, a longer one cut off.
    Text = "....."
    PokeStr StrPtr(Text), "ab"
    Assert.AreEqual "ab   ", Text

    Text = "....."
    PokeStr StrPtr(Text), "abcdefgh"
    Assert.AreEqual "abcde", Text

    Text = "..."
    PokeStr StrPtr(Text), ChrW$(&H20AC) & "1"
    Assert.AreEqual ChrW$(&H20AC) & "1 ", Text

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Poke")
Private Sub PokeObj_ReplacesObjectReference()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = New Collection
    Bag.Add "item"
    Dim Other As Collection
    Set Other = New Collection

    ' PokeObj takes the address of the object variable, not ObjPtr.
    Dim Holder As Collection
    '@Ignore ValueRequired
    PokeObj VarPtr(Holder), Bag
    Assert.IsTrue Holder Is Bag, "PokeObj into Nothing"

    Set Holder = Other
    '@Ignore ValueRequired
    PokeObj VarPtr(Holder), Bag
    Assert.IsTrue Holder Is Bag, "PokeObj replaces an object"
    Assert.AreEqual 0&, Other.Count

    ' The holder keeps its own reference.
    Set Bag = Nothing
    Assert.AreEqual "item", Holder.Item(1)

    '@Ignore ValueRequired
    PokeObj VarPtr(Holder), Nothing
    Assert.IsTrue Holder Is Nothing, "PokeObj Nothing"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: array pointers and properties
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Arrays")
Private Sub ArrVarPtr_ReturnsAddressOfArrayVariable()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.AreEqual RefVarPtrArray(Numbers), ArrVarPtr(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.AreEqual RefVarPtrArray(Numbers), ArrVarPtr(Numbers), "allocated"

    ' Not for a String array: a Declare call gets an ANSI copy of it.
    Dim Doubles() As Double
    ReDim Doubles(0 To 1)
    Assert.AreEqual RefVarPtrArray(Doubles), ArrVarPtr(Doubles), "Double array"

    ' For a Variant holding an array: the address of the data part.
    Dim Holder As Variant
    Holder = Array(1, 2)
    Assert.AreEqual VarPtr(Holder) + VariantDataOffset, ArrVarPtr(Holder), "Variant/array"

    Dim ErrNo As Long
    On Error Resume Next
    Dim Dummy As LongPtr: Dummy = ArrVarPtr(Holder(0))
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrTypeMismatch, ErrNo, "not an array"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrPtr_ReturnsSafeArrayAddress()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.AreEqual vbNullPtr, ArrPtr(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.AreEqual RefPeekPtr(RefVarPtrArray(Numbers)), ArrPtr(Numbers), "allocated"

    Dim Fixed(1 To 3) As Double
    Assert.AreEqual RefPeekPtr(RefVarPtrArray(Fixed)), ArrPtr(Fixed), "static"

    Dim Holder As Variant
    Holder = Array(1, 2)
    Assert.AreEqual RefPeekPtr(VarPtr(Holder) + VariantDataOffset), ArrPtr(Holder), "Variant/array"

    Dim ErrNo As Long
    On Error Resume Next
    Dim Dummy As LongPtr: Dummy = ArrPtr(42)
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrTypeMismatch, ErrNo, "not an array"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrDataPtr_ReturnsAddressOfFirstElement()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.AreEqual vbNullPtr, ArrDataPtr(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.AreEqual VarPtr(Numbers(1)), ArrDataPtr(Numbers), "Long array"

    Dim Grid() As Integer
    ReDim Grid(1 To 2, 0 To 3)
    Assert.AreEqual VarPtr(Grid(1, 0)), ArrDataPtr(Grid), "2-dimensional"

    Dim Words() As String
    ReDim Words(0 To 1)
    Assert.AreEqual VarPtr(Words(0)), ArrDataPtr(Words), "String array"

    Dim Mixed() As Variant
    ReDim Mixed(0 To 1)
    Assert.AreEqual VarPtr(Mixed(0)), ArrDataPtr(Mixed), "Variant array"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrDims_ReturnsNumberOfDimensions()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.AreEqual 0, ArrDims(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.AreEqual 1, ArrDims(Numbers), "1 dimension"
    ReDim Numbers(1 To 3, 0 To 1)
    Assert.AreEqual 2, ArrDims(Numbers), "2 dimensions"
    ReDim Numbers(1, 1, 1)
    Assert.AreEqual 3, ArrDims(Numbers), "3 dimensions"
    Erase Numbers
    Assert.AreEqual 0, ArrDims(Numbers), "erased"

    Assert.AreEqual 1, ArrDims(Array()), "Array()"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrElemSize_MatchesElementDistance()
    On Error GoTo TestFail

    Dim Bytes(0 To 1) As Byte
    Assert.AreEqual CLng(VarPtr(Bytes(1)) - VarPtr(Bytes(0))), ArrElemSize(Bytes), "Byte"
    Dim Flags(0 To 1) As Boolean
    Assert.AreEqual CLng(VarPtr(Flags(1)) - VarPtr(Flags(0))), ArrElemSize(Flags), "Boolean"
    Dim Shorts(0 To 1) As Integer
    Assert.AreEqual CLng(VarPtr(Shorts(1)) - VarPtr(Shorts(0))), ArrElemSize(Shorts), "Integer"
    Dim Longs(0 To 1) As Long
    Assert.AreEqual CLng(VarPtr(Longs(1)) - VarPtr(Longs(0))), ArrElemSize(Longs), "Long"
    Dim Doubles(0 To 1) As Double
    Assert.AreEqual CLng(VarPtr(Doubles(1)) - VarPtr(Doubles(0))), ArrElemSize(Doubles), "Double"
    Dim Amounts(0 To 1) As Currency
    Assert.AreEqual CLng(VarPtr(Amounts(1)) - VarPtr(Amounts(0))), ArrElemSize(Amounts), "Currency"
    Dim Words(0 To 1) As String
    Assert.AreEqual CLng(VarPtr(Words(1)) - VarPtr(Words(0))), ArrElemSize(Words), "String"
    Dim Things(0 To 1) As Object
    Assert.AreEqual PtrBytes, ArrElemSize(Things), "Object"
    Dim Mixed(0 To 1) As Variant
    Assert.AreEqual CLng(VarPtr(Mixed(1)) - VarPtr(Mixed(0))), ArrElemSize(Mixed), "Variant"

    Dim Numbers() As Long
    Assert.AreEqual 0&, ArrElemSize(Numbers), "unallocated"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub IsArrayStatic_TrueOnlyForFixedArrays()
    On Error GoTo TestFail

    Dim Fixed(1 To 3) As Long
    Assert.IsTrue IsArrayStatic(Fixed), "static"

    Dim Numbers() As Long
    Assert.IsFalse IsArrayStatic(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.IsFalse IsArrayStatic(Numbers), "dynamic"

    Assert.IsFalse IsArrayStatic(Array(1, 2)), "Array()"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub IsArrayAllocated_TrueWhenSafeArrayExists()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.IsFalse IsArrayAllocated(Numbers), "unallocated"
    ReDim Numbers(1 To 3)
    Assert.IsTrue IsArrayAllocated(Numbers), "ReDim"
    Erase Numbers
    Assert.IsFalse IsArrayAllocated(Numbers), "erased"

    Assert.IsTrue IsArrayAllocated(Split(vbNullString)), "Split of an empty string"
    Assert.IsFalse IsArrayAllocated(CreateArray(vbLong)), "CreateArray without dimensions"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub IsArrayEmpty_TrueWhenAnyDimensionHasNoElements()
    On Error GoTo TestFail

    Dim Numbers() As Long
    Assert.IsTrue IsArrayEmpty(Numbers), "unallocated"
    ReDim Numbers(1 To 1)
    Assert.IsFalse IsArrayEmpty(Numbers), "one element"
    ReDim Numbers(1 To 2, 1 To 3)
    Assert.IsFalse IsArrayEmpty(Numbers), "2x3"

    Assert.IsTrue IsArrayEmpty(Split(vbNullString)), "Split of an empty string"
    Assert.IsTrue IsArrayEmpty(CreateArray(vbLong, 2, 3, 0, 0, 0)), "3x0"
    Assert.IsTrue IsArrayEmpty(CreateArray(vbLong, 2, 0, 0, 3, 0)), "0x3"

    Dim Fixed(1 To 3) As Long
    Assert.IsFalse IsArrayEmpty(Fixed), "static"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: CreateArray, ArrElem and ArrBase
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Arrays")
Private Sub CreateArray_AllDimensionForms_GiveBounds()
    On Error GoTo TestFail

    Dim Created As Variant

    ' No dimensions: a typed, unallocated array.
    Created = CreateArray(vbLong)
    Assert.IsTrue VarType(Created) = vbArray + vbLong, "unallocated type"
    Assert.IsFalse IsArrayAllocated(Created), "unallocated"

    ' NDims and one number of elements for every dimension.
    Created = CreateArray(vbDouble, 2, 3)
    Assert.IsTrue VarType(Created) = vbArray + vbDouble, "NDims, count type"
    Assert.AreEqual "0..2 x 0..2", _
                    LBound(Created, 1) & ".." & UBound(Created, 1) & " x " & LBound(Created, 2) & ".." & UBound(Created, 2), _
                    "NDims, count"

    ' NDims and the number of elements per dimension.
    Created = CreateArray(vbString, 2, 3, 4)
    Assert.AreEqual "0..2 x 0..3", _
                    LBound(Created, 1) & ".." & UBound(Created, 1) & " x " & LBound(Created, 2) & ".." & UBound(Created, 2), _
                    "NDims, counts"
    Assert.AreEqual vbNullString, CStr(Created(2, 3))

    ' NDims and the number of elements and lower bound per dimension.
    Created = CreateArray(vbVariant, 2, 3, 1, 4, -2)
    Assert.AreEqual "1..3 x -2..1", _
                    LBound(Created, 1) & ".." & UBound(Created, 1) & " x " & LBound(Created, 2) & ".." & UBound(Created, 2), _
                    "NDims, counts, bounds"
    Assert.IsTrue IsEmpty(Created(1, -2)), "Variant elements are Empty"

    ' A SafeArray bound vector.
    Created = CreateArray(vbObject, Array(3, 1, 4, -2))
    Assert.AreEqual "1..3 x -2..1", _
                    LBound(Created, 1) & ".." & UBound(Created, 1) & " x " & LBound(Created, 2) & ".." & UBound(Created, 2), _
                    "bound vector"
    Assert.IsTrue Created(3, 1) Is Nothing, "Object elements are Nothing"

    ' One dimension with a lower bound.
    Created = CreateArray(vbByte, 1, 5, 10)
    Assert.AreEqual "10..14", LBound(Created) & ".." & UBound(Created), "1 dimension"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub CreateArray_InvalidArguments_Raise()
    On Error GoTo TestFail

    Dim Created As Variant
    Dim ErrNo As Long
    On Error Resume Next

    Created = CreateArray(vbEmpty, 1, 3)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrTypeMismatch, ErrNo, "vbEmpty"

    Created = CreateArray(vbDecimal, 1, 3)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrTypeMismatch, ErrNo, "vbDecimal"

    Created = CreateArray(vbUserDefinedType, 1, 3)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrTypeMismatch, ErrNo, "vbUserDefinedType"

    Created = CreateArray(vbLong, 0, 3)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "NDims = 0"

    Created = CreateArray(vbLong, 2)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrArgNotOptional, ErrNo, "NDims only"

    Created = CreateArray(vbLong, 2, 3, 4, 5)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "wrong number of arguments"

    Created = CreateArray(vbLong, Array(3, 0), 4)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "bound vector with more arguments"

    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrElemGet_ReadsElementByIndices()
    On Error GoTo TestFail

    Dim Grid() As Long
    ReDim Grid(1 To 3, 0 To 2)
    Grid(2, 1) = 21
    Grid(3, 0) = 30
    Dim Ix(0 To 1) As Long

    ' The indices are in VB order: Ix(0) is the first index.
    Ix(0) = 2: Ix(1) = 1
    Assert.AreEqual 21&, ArrElem(Grid, Ix)
    Ix(0) = 3: Ix(1) = 0
    Assert.AreEqual 30&, ArrElem(Grid, Ix)

    Dim Words() As String
    ReDim Words(0 To 2)
    Words(1) = "one"
    Dim Single1(0 To 0) As Long
    Single1(0) = 1
    Assert.AreEqual "one", ArrElem(Words, Single1)

    Dim Mixed As Variant
    Mixed = Array(1, 2.5, "x")
    Single1(0) = 1
    Assert.AreEqual 2.5, ArrElem(Mixed, Single1)

    Dim ErrNo As Long
    On Error Resume Next
    Ix(0) = 4: Ix(1) = 0
    Dim Dummy As Variant: Dummy = ArrElem(Grid, Ix)
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrOutOfRange, ErrNo, "index out of range"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrElemLet_WritesAndConvertsValue()
    On Error GoTo TestFail

    Dim Grid As Variant
    Grid = CreateArray(vbLong, 2, 3, 4)
    Dim Ix(0 To 1) As Long

    Ix(0) = 2: Ix(1) = 1
    ArrElem(Grid, Ix) = 77
    Assert.AreEqual 77&, CLng(Grid(2, 1))
    Assert.AreEqual 0&, CLng(Grid(1, 2))

    ' The value is converted to the array type.
    Ix(0) = 0: Ix(1) = 3
    ArrElem(Grid, Ix) = "12"
    Assert.AreEqual 12&, CLng(Grid(0, 3))

    Dim Words As Variant
    Words = CreateArray(vbString, 1, 2)
    Dim Single1(0 To 0) As Long
    Single1(0) = 1
    ArrElem(Words, Single1) = 42
    Assert.AreEqual "42", CStr(Words(1))

    Dim Mixed As Variant
    Mixed = Array(1, 2, 3)
    ArrElem(Mixed, Single1) = "two"
    Assert.AreEqual "two", Mixed(1)

    Dim ErrNo As Long
    On Error Resume Next
    ArrElem(Grid, Ix) = "abc"
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrTypeMismatch, ErrNo, "value cannot be converted"
    Ix(0) = 3
    ArrElem(Grid, Ix) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "index out of range"
    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrElemSet_SetsObjectElement()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = New Collection
    Dim Ix(0 To 1) As Long
    Ix(0) = 1: Ix(1) = 2

    Dim Things As Variant
    Things = CreateArray(vbObject, 2, 2, 3)
    Set ArrElem(Things, Ix) = Bag
    Assert.IsTrue Things(1, 2) Is Bag, "Object array"
    Assert.IsTrue Things(0, 0) Is Nothing, "other elements"

    Dim Mixed As Variant
    Mixed = CreateArray(vbVariant, 2, 2, 3)
    Set ArrElem(Mixed, Ix) = Bag
    Assert.IsTrue Mixed(1, 2) Is Bag, "Variant array"

    ' The array keeps its own reference.
    Set Bag = Nothing
    Assert.AreEqual 0&, Things(1, 2).Count

    Dim ErrNo As Long
    Dim Grid As Variant
    Grid = CreateArray(vbLong, 2, 2, 3)
    On Error Resume Next
    Set ArrElem(Grid, Ix) = New Collection
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrObjectRequired, ErrNo, "not an Object or Variant array"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Arrays")
Private Sub ArrBase_SetsLowerBoundOfDimension()
    On Error GoTo TestFail

    Dim Numbers() As Long
    ReDim Numbers(0 To 4)
    Numbers(0) = 10
    Numbers(4) = 14
    ArrBase(Numbers) = 1
    Assert.AreEqual "1..5", LBound(Numbers) & ".." & UBound(Numbers)
    Assert.AreEqual 10&, Numbers(1)
    Assert.AreEqual 14&, Numbers(5)

    Dim Grid() As Long
    ReDim Grid(0 To 2, 0 To 3)
    ArrBase(Grid, 2) = -1
    Assert.AreEqual "0..2 x -1..2", _
                    LBound(Grid, 1) & ".." & UBound(Grid, 1) & " x " & LBound(Grid, 2) & ".." & UBound(Grid, 2)
    ArrBase(Grid, 1) = 5
    Assert.AreEqual "5..7 x -1..2", _
                    LBound(Grid, 1) & ".." & UBound(Grid, 1) & " x " & LBound(Grid, 2) & ".." & UBound(Grid, 2)

    Dim ErrNo As Long
    On Error Resume Next
    ArrBase(Grid, 3) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "dimension too high"
    ArrBase(Grid, 0) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "dimension 0"
    Dim Unallocated() As Long
    ArrBase(Unallocated) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "unallocated"
    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: MidArr and CloneArray
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Slices")
Private Sub MidArrGet_ReturnsSlice()
    On Error GoTo TestFail

    Dim Numbers() As Long
    ReDim Numbers(1 To 10)
    Dim Position As Long
    For Position = 1 To 10
        Numbers(Position) = 100 + Position
    Next

    ' The slice keeps the lower bound of the array.
    Dim Slice As Variant
    Slice = MidArr(Numbers, 3, 4)
    Assert.IsTrue VarType(Slice) = vbArray + vbLong, "VarType"
    Assert.AreEqual "1..4", LBound(Slice) & ".." & UBound(Slice), "slice bounds"
    Assert.AreEqual "103,104,105,106", ListOf(Slice), "slice"

    Slice = MidArr(Numbers, 8)
    Assert.AreEqual "108,109,110", ListOf(Slice), "no length: up to the end"

    Slice = MidArr(Numbers, 9, 5)
    Assert.AreEqual "109,110", ListOf(Slice), "length cut off at the end"

    Slice = MidArr(Numbers, 1, 0)
    Assert.IsTrue VarType(Slice) = vbArray + vbLong, "length 0 type"
    Assert.IsFalse IsArrayAllocated(Slice), "length 0"

    Slice = MidArr(Numbers, 11)
    Assert.IsTrue VarType(Slice) = vbArray + vbLong, "start after the end type"
    Assert.IsFalse IsArrayAllocated(Slice), "start after the end"

    Dim Unallocated() As Long
    Slice = MidArr(Unallocated, 1)
    Assert.IsFalse IsArrayAllocated(Slice), "unallocated"

    ' The array itself is unchanged.
    Assert.AreEqual "1..10", LBound(Numbers) & ".." & UBound(Numbers), "array bounds"
    Assert.AreEqual 101&, Numbers(1)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Slices")
Private Sub MidArrGet_ReferenceTypes_AreCopied()
    On Error GoTo TestFail

    Dim Words() As String
    Words = Split("a b c d e")
    Dim Slice As Variant
    Slice = MidArr(Words, 1, 3)
    Assert.AreEqual "b,c,d", Join(Slice, ",")
    Slice(0) = "changed"
    Assert.AreEqual "a b c d e", Join(Words, " "), "String slice is a copy"

    Dim Things(0 To 2) As Object
    Set Things(0) = New Collection
    Set Things(1) = New Collection
    Set Things(2) = New Collection
    Slice = MidArr(Things, 1)
    Assert.IsTrue Slice(0) Is Things(1), "Object slice"
    Assert.IsTrue Slice(1) Is Things(2), "Object slice"

    Dim Mixed As Variant
    Mixed = Array(1, "two", 3.5, Things(0))
    Slice = MidArr(Mixed, 1, 3)
    Assert.AreEqual "two", Slice(0)
    Assert.IsTrue Slice(2) Is Things(0), "Variant slice with object"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Slices")
Private Sub MidArrGet_InvalidArguments_Raise()
    On Error GoTo TestFail

    Dim Numbers() As Long
    ReDim Numbers(1 To 10)
    Dim Grid() As Long
    ReDim Grid(1 To 2, 1 To 2)
    Dim Slice As Variant
    Dim ErrNo As Long
    On Error Resume Next

    Slice = MidArr(Numbers, 0)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "start before the lower bound"

    Slice = MidArr(Numbers, 2, -1)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "negative length"

    Slice = MidArr(Grid, 1)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "2 dimensions"

    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Slices")
Private Sub MidArrLet_ReplacesSlice()
    On Error GoTo TestFail

    Dim Numbers() As Long
    ReDim Numbers(1 To 8)
    Dim Patch() As Long
    ReDim Patch(0 To 2)
    Patch(0) = 7: Patch(1) = 8: Patch(2) = 9

    MidArr(Numbers, 3) = Patch
    Assert.AreEqual "0,0,7,8,9,0,0,0", ListOf(Numbers)

    ReDim Numbers(1 To 8)
    MidArr(Numbers, 3, 2) = Patch
    Assert.AreEqual "0,0,7,8,0,0,0,0", ListOf(Numbers), "length"

    ReDim Numbers(1 To 8)
    MidArr(Numbers, 7) = Patch
    Assert.AreEqual "0,0,0,0,0,0,7,8", ListOf(Numbers), "cut off at the end"

    ReDim Numbers(1 To 8)
    MidArr(Numbers, 3, 0) = Patch
    MidArr(Numbers, 9) = Patch
    Assert.AreEqual "0,0,0,0,0,0,0,0", ListOf(Numbers), "length 0 and start after the end"

    Assert.AreEqual "7,8,9", ListOf(Patch), "patch unchanged"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Slices")
Private Sub MidArrLet_ReferenceTypes_KeepSourceIntact()
    On Error GoTo TestFail

    Dim Words() As String
    Words = Split("a b c d e")
    Dim Patch() As String
    Patch = Split("x y")
    MidArr(Words, 2) = Patch
    Assert.AreEqual "a b x y e", Join(Words, " ")
    Assert.AreEqual "x y", Join(Patch, " "), "String source intact"
    Patch(0) = "changed"
    Assert.AreEqual "a b x y e", Join(Words, " "), "String target is a copy"

    Dim Bag As Collection
    Set Bag = New Collection
    Dim Things(0 To 2) As Object
    Dim Parts() As Object
    ReDim Parts(0 To 0)
    Set Parts(0) = Bag
    MidArr(Things, 1) = Parts
    Assert.IsTrue Things(1) Is Bag, "Object target"
    Assert.IsTrue Parts(0) Is Bag, "Object source intact"
    Assert.IsTrue Things(0) Is Nothing, "Object neighbour"

    Dim Mixed As Variant
    Mixed = Array(1, 2, 3, 4)
    MidArr(Mixed, 1) = Array("two", Bag)
    Assert.AreEqual "two", Mixed(1)
    Assert.IsTrue Mixed(2) Is Bag, "Variant target with object"
    Assert.AreEqual 4, Mixed(3)

    Dim ErrNo As Long
    On Error Resume Next
    MidArr(Words, -1) = Patch
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "start before the lower bound"
    MidArr(Words, 1) = Array("v")
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrTypeMismatch, ErrNo, "types differ"
    On Error GoTo TestFail
    Assert.AreEqual "a b x y e", Join(Words, " "), "unchanged after the errors"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Slices")
Private Sub CloneArray_ReturnsIndependentCopy()
    On Error GoTo TestFail

    Dim Grid() As Long
    ReDim Grid(1 To 2, 0 To 2)
    Grid(1, 0) = 10
    Grid(2, 2) = 22
    Dim Clone As Variant
    Clone = CloneArray(ArrPtr(Grid))
    Assert.IsTrue VarType(Clone) = vbArray + vbLong, "VarType"
    Assert.AreEqual "1..2 x 0..2", _
                    LBound(Clone, 1) & ".." & UBound(Clone, 1) & " x " & LBound(Clone, 2) & ".." & UBound(Clone, 2)
    Assert.AreEqual 10&, CLng(Clone(1, 0))
    Assert.AreEqual 22&, CLng(Clone(2, 2))
    Clone(1, 0) = 99
    Assert.AreEqual 10&, Grid(1, 0), "Long clone is a copy"
    Assert.IsTrue ArrPtr(Clone) <> ArrPtr(Grid), "own SafeArray"

    Dim Words() As String
    Words = Split("a b c")
    Clone = CloneArray(ArrPtr(Words))
    Assert.IsTrue VarType(Clone) = vbArray + vbString, "VarType"
    Clone(0) = "changed"
    Assert.AreEqual "a b c", Join(Words, " "), "String clone is a copy"

    Dim Mixed As Variant
    Mixed = Array(1, "two")
    Clone = CloneArray(ArrPtr(Mixed))
    Assert.IsTrue VarType(Clone) = vbArray + vbVariant, "VarType"
    Assert.AreEqual "two", Clone(1)

    Dim ErrNo As Long
    On Error Resume Next
    Clone = CloneArray(vbNullPtr)
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrInvalidCall, ErrNo, "null address"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: memory blocks
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Memory")
Private Sub CopyMem_AllSmallSizes_CopyExactBlock()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Size As Long
    For Size = 1 To 64
        Dim Original() As Byte
        ReDim Original(0 To 79)
        Dim Position As Long
        For Position = 0 To 79
            Original(Position) = (37 * Position + 11) Mod 256
        Next
        Dim Buffer() As Byte: Buffer = FilledBuffer(80)
        CopyMem VarPtr(Buffer(8)), VarPtr(Original(3)), Size

        Dim Correct As Boolean: Correct = FillerOutside(Buffer, 8, Size) = 80 - Size
        For Position = 0 To Size - 1
            If Buffer(8 + Position) <> Original(3 + Position) Then Correct = False
        Next
        If Not Correct Then
            Failures = Failures + 1
            If Failures = 1 Then FirstFailure = "size " & Size
        End If
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Memory")
Private Sub CopyMem_OverlappingBlocks_CopyLikeMemmove()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Size As Long
    For Size = 1 To 40
        Dim Shift As Long
        For Shift = -3 To 3
            If Shift <> 0 Then
                Dim Block() As Byte
                ReDim Block(0 To 59)
                Dim Position As Long
                For Position = 0 To 59
                    Block(Position) = Position
                Next
                ' Expected: the bytes 10..10+Size-1 moved to 10+Shift.
                Dim Expected() As Byte: Expected = Block
                For Position = 0 To Size - 1
                    Expected(10 + Shift + Position) = 10 + Position
                Next
                CopyMem VarPtr(Block(10 + Shift)), VarPtr(Block(10)), Size
                If StrConv(Block, vbUnicode) <> StrConv(Expected, vbUnicode) Then
                    Failures = Failures + 1
                    If Failures = 1 Then FirstFailure = "size " & Size & " shift " & Shift
                End If
            End If
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Memory")
Private Sub CopyMem_LargeBlockAndEdgeCases()
    On Error GoTo TestFail

    ' Above the threshold the API is used.
    Const LargeSize As Long = 1000001
    Dim Original() As Byte
    ReDim Original(0 To LargeSize - 1)
    Original(0) = 1
    Original(LargeSize \ 2) = 2
    Original(LargeSize - 1) = 3
    Dim Buffer() As Byte
    ReDim Buffer(0 To LargeSize)
    Buffer(LargeSize) = Filler
    CopyMem VarPtr(Buffer(0)), VarPtr(Original(0)), LargeSize
    Assert.AreEqual "1 2 3 165", _
                    Buffer(0) & " " & Buffer(LargeSize \ 2) & " " & Buffer(LargeSize - 1) & " " & Buffer(LargeSize), _
                    "large block"

    ' Zero or a negative number of bytes does nothing.
    Buffer = FilledBuffer(8)
    CopyMem VarPtr(Buffer(0)), VarPtr(Original(0)), 0
    CopyMem VarPtr(Buffer(0)), VarPtr(Original(0)), -1
    Assert.AreEqual 8&, FillerOutside(Buffer, 0, 0), "no bytes"

    Dim ErrNo As Long
    On Error Resume Next
    CopyMem vbNullPtr, VarPtr(Original(0)), 4
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "null destination"
    CopyMem VarPtr(Buffer(0)), vbNullPtr, 4
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "null source"
    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Memory")
Private Sub ZeroMem_AllSizes_ZeroExactBlock()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Size As Variant
    For Each Size In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17, 31, 32, 33, 63, 64, 9999, 10000, 10001)
        Dim Buffer() As Byte: Buffer = FilledBuffer(Size + 16)
        ZeroMem VarPtr(Buffer(8)), Size
        Dim Correct As Boolean: Correct = FillerOutside(Buffer, 8, Size) = 16
        Dim Position As Long
        For Position = 8 To 8 + Size - 1
            If Buffer(Position) <> 0 Then Correct = False
        Next
        If Not Correct Then
            Failures = Failures + 1
            If Failures = 1 Then FirstFailure = "size " & Size
        End If
    Next
    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

    Buffer = FilledBuffer(8)
    ZeroMem VarPtr(Buffer(0)), 0
    ZeroMem VarPtr(Buffer(0)), -1
    Assert.AreEqual 8&, FillerOutside(Buffer, 0, 0), "no bytes"

    Dim ErrNo As Long
    On Error Resume Next
    ZeroMem vbNullPtr, 4
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrInvalidCall, ErrNo, "null address"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Memory")
Private Sub DumpMem_ShowsAddressBytesAndText()
    On Error GoTo TestFail

    Dim Block() As Byte
    ReDim Block(0 To 19)
    Dim Position As Long
    For Position = 0 To 19
        Block(Position) = Position
    Next
    Block(16) = Asc("A"): Block(17) = Asc("b"): Block(18) = Asc("C")

    Dim Dump As String
    Dump = DumpMem(VarPtr(Block(0)), 20)
    Dim Lines() As String
    Lines = Split(Replace(Dump, vbLf, vbNullString), vbCr)

    Dim FirstAddress As String
    FirstAddress = Right$(String$(AddressHexDigits, "0") & Hex$(VarPtr(Block(0))), AddressHexDigits)
    Dim SecondAddress As String
    SecondAddress = Right$(String$(AddressHexDigits, "0") & Hex$(VarPtr(Block(0)) + 16), AddressHexDigits)

    ' Header, title, two lines of bytes and the footer.
    Assert.IsTrue Lines(0) Like "=*=", "header"
    Assert.IsTrue InStr(Lines(1), Hex$(VarPtr(Block(0)))) > 0, "title shows the address"
    Assert.IsTrue Lines(2) Like FirstAddress & " 0000  00 01 02 03 04 05 06 07-08 09 0A 0B 0C 0D 0E 0F  ................", _
                  "first line: " & Lines(2)
    Assert.IsTrue Lines(3) Like SecondAddress & " 0010  41 62 43 13 *AbC.*", "second line: " & Lines(3)
    Assert.IsTrue Lines(4) Like "=*=", "footer"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Memory")
Private Sub DumpMem_LineWidthAndEdgeCases()
    On Error GoTo TestFail

    Dim Block() As Byte
    ReDim Block(0 To 15)
    Dim Position As Long
    For Position = 0 To 15
        Block(Position) = Position
    Next

    ' A line width below 8 is raised to 8.
    Dim Lines() As String
    Lines = Split(Replace(DumpMem(VarPtr(Block(0)), 16, 4, 4), vbLf, vbNullString), vbCr)
    Assert.IsTrue Lines(2) Like "* 0000  00 01 02 03 04 05 06 07  ........", "first line: " & Lines(2)
    Assert.IsTrue Lines(3) Like "* 0008  08 09 0A 0B 0C 0D 0E 0F  ........", "second line: " & Lines(3)

    Assert.AreEqual vbNullString, DumpMem(VarPtr(Block(0)), 0), "no bytes"

    Dim ErrNo As Long
    On Error Resume Next
    Dim Dump As String: Dump = DumpMem(vbNullPtr, 16)
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrInvalidCall, ErrNo, "null address"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: collections
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Collections")
Private Sub CollKeys_ReturnsKeysInOrder()
    On Error GoTo TestFail

    Dim Keys() As String
    Keys = CollKeys(SampleBag())
    Assert.AreEqual "0..4", LBound(Keys) & ".." & UBound(Keys)
    Assert.AreEqual "A,,C,D,E", Join(Keys, ",")

    Keys = CollKeys(SampleBag(), 1)
    Assert.AreEqual "1..5", LBound(Keys) & ".." & UBound(Keys), "base 1"

    Keys = CollKeys(New Collection)
    Assert.IsFalse IsArrayAllocated(Keys), "empty collection"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collections")
Private Sub CollKey_ReturnsKeyAtPosition()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = SampleBag()
    ' Positions 1-2 are found from the head, 3-5 from the tail.
    Assert.AreEqual "A", CollKey(Bag, 1)
    Assert.AreEqual vbNullString, CollKey(Bag, 2), "item without a key"
    Assert.AreEqual "C", CollKey(Bag, 3)
    Assert.AreEqual "D", CollKey(Bag, 4)
    Assert.AreEqual "E", CollKey(Bag, 5)

    Dim ErrNo As Long
    Dim KeyText As String
    On Error Resume Next
    KeyText = CollKey(Bag, 0)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "position 0"
    KeyText = CollKey(Bag, 6)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "position after the end"
    KeyText = CollKey(New Collection, 1)
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "empty collection"
    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collections")
Private Sub CollPos_ReturnsPositionOfKey()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = SampleBag()
    Assert.AreEqual 1&, CollPos(Bag, "A")
    Assert.AreEqual 3&, CollPos(Bag, "C")
    Assert.AreEqual 5&, CollPos(Bag, "E")
    Assert.AreEqual 4&, CollPos(Bag, "d"), "keys are case-insensitive"
    Assert.AreEqual 0&, CollPos(Bag, "Q"), "missing key"

    Dim ErrNo As Long
    Dim Found As Long
    On Error Resume Next
    Found = CollPos(New Collection, "A")
    ErrNo = Err.Number
    On Error GoTo TestFail
    Assert.AreEqual ErrInvalidCall, ErrNo, "empty collection"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collections")
Private Sub CollItem_ReplacesItemKeepingKeyAndOrder()
    On Error GoTo TestFail

    Dim Bag As Collection
    Set Bag = SampleBag()
    CollItem(Bag, 1) = "uno"
    CollItem(Bag, 4) = 40
    CollItem(Bag, "e") = "fin"
    Dim Other As Collection
    Set Other = New Collection
    '@Ignore ValueRequired
    CollItem(Bag, 2) = Other

    Assert.AreEqual 5&, Bag.Count
    Assert.AreEqual "uno", Bag.Item("A"), "by position, from the head"
    Assert.IsTrue Bag.Item(2) Is Other, "object item"
    Assert.AreEqual "three", Bag.Item(3), "untouched item"
    Assert.AreEqual 40, Bag.Item("D"), "by position, from the tail"
    Assert.AreEqual "fin", Bag.Item(5), "by key"
    Assert.AreEqual "A,,C,D,E", Join(CollKeys(Bag), ","), "keys unchanged"

    Dim ErrNo As Long
    On Error Resume Next
    CollItem(Bag, 6) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrOutOfRange, ErrNo, "position after the end"
    CollItem(Bag, "Q") = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "missing key"
    CollItem(New Collection, 1) = 1
    ErrNo = Err.Number: Err.Clear
    Assert.AreEqual ErrInvalidCall, ErrNo, "empty collection"
    On Error GoTo TestFail

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
