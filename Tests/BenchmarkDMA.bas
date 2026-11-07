Attribute VB_Name = "BenchmarkDMA"
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, FunctionReturnValueDiscarded, VariableNotUsed, AssignmentNotUsed, IntegerDataType
'@Folder("Tests")
'@ModuleDescription("Benchmarks for DMA; the results go to the Immediate window.")
Option Compare Database
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Benchmarks for DMA
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' Run from the Immediate window:
'     RunBenchmarks             all tables (about half a minute)
'     RunBenchmarks 0.1         a quick run with a tenth of the repetitions
'     BenchmarkCopyMem          one table (every table has its own method)
'
' The times are in 탎 per call, after subtracting the time of an empty loop.
' The API columns use RtlMoveMemory and RtlZeroMemory. The LSet columns use the
' private LSetCopyMem and LSetZeroMem: CopyMem and ZeroMem without the switch to
' the API for large blocks. The benchmarks are not Rubberduck tests, so the Test
' Explorer stays fast. Needs the Stopwatch class.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' https://learn.microsoft.com/en-us/windows/win32/devnotes/rtlmovememory
Private Declare PtrSafe Sub ApiCopy Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal targetPtr As LongPtr, _
    ByVal sourcePtr As LongPtr, _
    ByVal byteCount As Long _
)

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlzeromemory
Private Declare PtrSafe Sub ApiZero Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByVal targetPtr As LongPtr, _
    ByVal byteCount As Long _
)

#If Win64 Then
Private Const PtrBytes As Long = 8
#Else
Private Const PtrBytes As Long = 4
#End If

' The number of calls for the fast methods at a work factor of 1.
Private Const FastCalls As Long = 1000000

' The timed Peek and Poke methods; pmBaseline is the empty loop.
Private Enum PeekPokeMethod
    pmBaseline
    pmPeekBool
    pmApiPeekBool
    pmPokeBool
    pmApiPokeBool
    pmPeekByte
    pmApiPeekByte
    pmPokeByte
    pmApiPokeByte
    pmPeekCur
    pmApiPeekCur
    pmPokeCur
    pmApiPokeCur
    pmPeekDbl
    pmApiPeekDbl
    pmPokeDbl
    pmApiPokeDbl
    pmPeekInt
    pmApiPeekInt
    pmPokeInt
    pmApiPokeInt
    pmPeekLng
    pmApiPeekLng
    pmPokeLng
    pmApiPokeLng
    pmPeekLngLng
    pmApiPeekLngLng
    pmPokeLngLng
    pmApiPokeLngLng
    pmPeekPtr
    pmApiPeekPtr
    pmPokePtr
    pmApiPokePtr
    pmPeekSng
    pmApiPeekSng
    pmPokeSng
    pmApiPokeSng
    pmPeekVar
    pmPokeVar
    pmPeekObj
    pmPokeObj
End Enum

' The timed array methods; amBaseline is the empty loop.
Private Enum ArrayMethod
    amBaseline
    amArrPtr
    amArrVarPtr
    amArrDataPtr
    amArrDims
    amIsArrayEmpty
    amLast = amIsArrayEmpty
End Enum

' The timed string methods; tmBaseline is the empty loop.
Private Enum StringMethod
    tmBaseline
    tmPeekStr
    tmPokeStr
    tmAssign
End Enum

' The timed slice methods.
Private Enum SliceMethod
    smGetMidArr
    smGetLoop
    smLetMidArr
    smLetLoop
End Enum

' The timed memory block methods.
Private Enum BlockMethod
    bmCopyMem
    bmLSetCopy
    bmApiCopy
    bmZeroMem
    bmLSetZero
    bmApiZero
End Enum

' The timed collection methods.
Private Enum CollectionMethod
    cmItemByPosition
    cmItemByKey
    cmCollKey
    cmCollPos
    cmCollKeys
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Prints every DMA benchmark table; a work factor below 1 gives a quicker run.")
'@Ignore ProcedureNotUsed
Public Sub RunBenchmarks(Optional ByVal workFactor As Double = 1)

    BenchmarkPeekPoke workFactor
    BenchmarkStrings workFactor
    BenchmarkArrays workFactor
    BenchmarkMidArr workFactor
    BenchmarkCopyMem workFactor
    BenchmarkZeroMem workFactor
    BenchmarkCollections workFactor

End Sub


'@Description("Prints the time per call of every Peek and Poke method and of the API.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkPeekPoke(Optional ByVal workFactor As Double = 1)

    Dim Calls As Long: Calls = Repeats(workFactor, FastCalls)
    Dim LoopBase As Double: LoopBase = PeekPokeTime(pmBaseline, Calls)

    Debug.Print
    Debug.Print "Peek and Poke, " & Format$(Calls, "#,##0") & " calls (탎 per call)"
    Debug.Print "Method", "DMA", "API"
    PrintPair "PeekBool", PeekPokeTime(pmPeekBool, Calls), PeekPokeTime(pmApiPeekBool, Calls), LoopBase, Calls
    PrintPair "PokeBool", PeekPokeTime(pmPokeBool, Calls), PeekPokeTime(pmApiPokeBool, Calls), LoopBase, Calls
    PrintPair "PeekByte", PeekPokeTime(pmPeekByte, Calls), PeekPokeTime(pmApiPeekByte, Calls), LoopBase, Calls
    PrintPair "PokeByte", PeekPokeTime(pmPokeByte, Calls), PeekPokeTime(pmApiPokeByte, Calls), LoopBase, Calls
    PrintPair "PeekCur", PeekPokeTime(pmPeekCur, Calls), PeekPokeTime(pmApiPeekCur, Calls), LoopBase, Calls
    PrintPair "PokeCur", PeekPokeTime(pmPokeCur, Calls), PeekPokeTime(pmApiPokeCur, Calls), LoopBase, Calls
    PrintPair "PeekDbl", PeekPokeTime(pmPeekDbl, Calls), PeekPokeTime(pmApiPeekDbl, Calls), LoopBase, Calls
    PrintPair "PokeDbl", PeekPokeTime(pmPokeDbl, Calls), PeekPokeTime(pmApiPokeDbl, Calls), LoopBase, Calls
    PrintPair "PeekInt", PeekPokeTime(pmPeekInt, Calls), PeekPokeTime(pmApiPeekInt, Calls), LoopBase, Calls
    PrintPair "PokeInt", PeekPokeTime(pmPokeInt, Calls), PeekPokeTime(pmApiPokeInt, Calls), LoopBase, Calls
    PrintPair "PeekLng", PeekPokeTime(pmPeekLng, Calls), PeekPokeTime(pmApiPeekLng, Calls), LoopBase, Calls
    PrintPair "PokeLng", PeekPokeTime(pmPokeLng, Calls), PeekPokeTime(pmApiPokeLng, Calls), LoopBase, Calls
#If Win64 Then
    PrintPair "PeekLngLng", PeekPokeTime(pmPeekLngLng, Calls), PeekPokeTime(pmApiPeekLngLng, Calls), LoopBase, Calls
    PrintPair "PokeLngLng", PeekPokeTime(pmPokeLngLng, Calls), PeekPokeTime(pmApiPokeLngLng, Calls), LoopBase, Calls
#End If
    PrintPair "PeekPtr", PeekPokeTime(pmPeekPtr, Calls), PeekPokeTime(pmApiPeekPtr, Calls), LoopBase, Calls
    PrintPair "PokePtr", PeekPokeTime(pmPokePtr, Calls), PeekPokeTime(pmApiPokePtr, Calls), LoopBase, Calls
    PrintPair "PeekSng", PeekPokeTime(pmPeekSng, Calls), PeekPokeTime(pmApiPeekSng, Calls), LoopBase, Calls
    PrintPair "PokeSng", PeekPokeTime(pmPokeSng, Calls), PeekPokeTime(pmApiPokeSng, Calls), LoopBase, Calls
    Debug.Print "PeekVar", Format$(Microseconds(PeekPokeTime(pmPeekVar, Calls) - LoopBase, Calls), "0.000")
    Debug.Print "PokeVar", Format$(Microseconds(PeekPokeTime(pmPokeVar, Calls) - LoopBase, Calls), "0.000")
    Debug.Print "PeekObj", Format$(Microseconds(PeekPokeTime(pmPeekObj, Calls) - LoopBase, Calls), "0.000")
    Debug.Print "PokeObj", Format$(Microseconds(PeekPokeTime(pmPokeObj, Calls) - LoopBase, Calls), "0.000")
End Sub


'@Description("Prints the time per call of PeekStr, PokeStr and a VBA assignment by string length.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkStrings(Optional ByVal workFactor As Double = 1)

    Debug.Print
    Debug.Print "PeekStr and PokeStr by string length (탎 per call)"
    Debug.Print "Chars", "PeekStr", "PokeStr", "VBA a = b"
    Dim TextLength As Variant
    For Each TextLength In Array(10, 100, 1000, 10000)
        Dim Calls As Long: Calls = Repeats(workFactor, FastCalls \ (1 + TextLength \ 10))
        Dim LoopBase As Double: LoopBase = StringTime(tmBaseline, TextLength, Calls)
        Debug.Print Format$(TextLength, "#,##0"), _
                    Format$(Microseconds(StringTime(tmPeekStr, TextLength, Calls) - LoopBase, Calls), "0.000"), _
                    Format$(Microseconds(StringTime(tmPokeStr, TextLength, Calls) - LoopBase, Calls), "0.000"), _
                    Format$(Microseconds(StringTime(tmAssign, TextLength, Calls) - LoopBase, Calls), "0.000")
    Next

End Sub


'@Description("Prints the time per call of the array pointer and property methods.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkArrays(Optional ByVal workFactor As Double = 1)

    Dim Calls As Long: Calls = Repeats(workFactor, FastCalls)

    Dim Longs() As Long
    ReDim Longs(1 To 100)
    Dim NoLongs() As Long
    Dim Mixed() As Variant
    ReDim Mixed(1 To 100)
    Dim NoMixed() As Variant

    Dim LongsBase As Double: LongsBase = ArrayTime(amBaseline, Longs, Calls)
    Dim NoLongsBase As Double: NoLongsBase = ArrayTime(amBaseline, NoLongs, Calls)
    Dim MixedBase As Double: MixedBase = ArrayTime(amBaseline, Mixed, Calls)
    Dim NoMixedBase As Double: NoMixedBase = ArrayTime(amBaseline, NoMixed, Calls)

    Debug.Print
    Debug.Print "Array methods, " & Format$(Calls, "#,##0") & " calls (탎 per call)"
    Debug.Print "Method", "Long(100)", "Long()", "Variant(100)", "Variant()"
    Dim Method As ArrayMethod
    For Method = amArrPtr To amLast
        Debug.Print Split("Baseline ArrPtr ArrVarPtr ArrDataPtr ArrDims IsArrayEmpty")(Method), _
                    Format$(Microseconds(ArrayTime(Method, Longs, Calls) - LongsBase, Calls), "0.000"), _
                    Format$(Microseconds(ArrayTime(Method, NoLongs, Calls) - NoLongsBase, Calls), "0.000"), _
                    Format$(Microseconds(ArrayTime(Method, Mixed, Calls) - MixedBase, Calls), "0.000"), _
                    Format$(Microseconds(ArrayTime(Method, NoMixed, Calls) - NoMixedBase, Calls), "0.000")
    Next

End Sub


'@Description("Prints the time per call of MidArr and of a For loop for Long arrays.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkMidArr(Optional ByVal workFactor As Double = 1)

    Debug.Print
    Debug.Print "MidArr for Long(1 To Length) (탎 per call)"
    Debug.Print "Length", "Get MidArr", "For loop", "Let MidArr", "For loop"
    Dim Length As Variant
    For Each Length In Array(1, 10, 100, 1000, 10000, 100000)
        Dim Calls As Long: Calls = Repeats(workFactor, FastCalls \ (Length + 100))
        Debug.Print Format$(Length, "#,##0"), _
                    Format$(Microseconds(SliceTime(smGetMidArr, Length, Calls), Calls), "0.000"), _
                    Format$(Microseconds(SliceTime(smGetLoop, Length, Calls), Calls), "0.000"), _
                    Format$(Microseconds(SliceTime(smLetMidArr, Length, Calls), Calls), "0.000"), _
                    Format$(Microseconds(SliceTime(smLetLoop, Length, Calls), Calls), "0.000")
    Next

End Sub


'@Description("Prints the time per call of CopyMem, of its LSet method alone and of the CopyMemory API.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkCopyMem(Optional ByVal workFactor As Double = 1)

    Debug.Print
    Debug.Print "CopyMem, its LSet method alone and the CopyMemory API (탎 per call)"
    Debug.Print "NBytes", "CopyMem", "LSet only", "CopyMemory"
    Dim NBytes As Variant
    For Each NBytes In Array(10, 100, 1000, 10000, 100000, 1000000, 10000000)
        Dim Calls As Long: Calls = Repeats(workFactor, 500000000 \ (NBytes + 1000))
        Debug.Print Format$(NBytes, "#,##0"), _
                    Format$(Microseconds(BlockTime(bmCopyMem, NBytes, Calls), Calls), "0.000"), _
                    Format$(Microseconds(BlockTime(bmLSetCopy, NBytes, Calls), Calls), "0.000"), _
                    Format$(Microseconds(BlockTime(bmApiCopy, NBytes, Calls), Calls), "0.000")
    Next

End Sub


'@Description("Prints the time per call of ZeroMem, of its LSet method alone and of the ZeroMemory API.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkZeroMem(Optional ByVal workFactor As Double = 1)

    Debug.Print
    Debug.Print "ZeroMem, its LSet method alone and the ZeroMemory API (탎 per call)"
    Debug.Print "NBytes", "ZeroMem", "LSet only", "ZeroMemory"
    Dim NBytes As Variant
    For Each NBytes In Array(10, 100, 1000, 10000, 100000, 1000000, 10000000)
        Dim Calls As Long: Calls = Repeats(workFactor, 500000000 \ (NBytes + 1000))
        Debug.Print Format$(NBytes, "#,##0"), _
                    Format$(Microseconds(BlockTime(bmZeroMem, NBytes, Calls), Calls), "0.000"), _
                    Format$(Microseconds(BlockTime(bmLSetZero, NBytes, Calls), Calls), "0.000"), _
                    Format$(Microseconds(BlockTime(bmApiZero, NBytes, Calls), Calls), "0.000")
    Next

End Sub


'@Description("Prints the time per call of Item and of the Collection extensions.")
'@Ignore ProcedureNotUsed
Public Sub BenchmarkCollections(Optional ByVal workFactor As Double = 1)

    Debug.Print
    Debug.Print "Collection with keyed items, item at position Count \ 2 (탎 per call)"
    Debug.Print "Count", ".Item pos", ".Item key", "CollKey", "CollPos", "CollKeys"
    Dim ItemCount As Variant
    For Each ItemCount In Array(10, 100, 1000, 10000, 100000)
        Dim Bag As Collection
        Set Bag = New Collection
        Dim Position As Long
        For Position = 1 To ItemCount
            Bag.Add Position, "K" & Position
        Next
        Dim Calls As Long: Calls = Repeats(workFactor, 2 * FastCalls \ (ItemCount + 100))
        Debug.Print Format$(ItemCount, "#,##0"), _
                    Format$(Microseconds(CollectionTime(cmItemByPosition, Bag, Calls), Calls), "0.000"), _
                    Format$(Microseconds(CollectionTime(cmItemByKey, Bag, Calls), Calls), "0.000"), _
                    Format$(Microseconds(CollectionTime(cmCollKey, Bag, Calls), Calls), "0.000"), _
                    Format$(Microseconds(CollectionTime(cmCollPos, Bag, Calls), Calls), "0.000"), _
                    Format$(Microseconds(CollectionTime(cmCollKeys, Bag, Calls), Calls), "0.000")
    Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the number of calls for a work factor, at least 3.")
Private Function Repeats(ByVal workFactor As Double, ByVal baseCount As Long) As Long

    Repeats = CLng(workFactor * baseCount)
    If Repeats < 3 Then Repeats = 3

End Function


'@Description("Returns the time per call in 탎.")
Private Function Microseconds(ByVal seconds As Double, ByVal callCount As Long) As Double

    Microseconds = seconds * 1000000 / callCount

End Function


'@Description("Prints one table row with a DMA time and an API time per call.")
Private Sub PrintPair( _
    ByVal methodName As String, _
    ByVal dmaSeconds As Double, _
    ByVal apiSeconds As Double, _
    ByVal baseSeconds As Double, _
    ByVal callCount As Long _
)

    Debug.Print methodName, _
                Format$(Microseconds(dmaSeconds - baseSeconds, callCount), "0.000"), _
                Format$(Microseconds(apiSeconds - baseSeconds, callCount), "0.000")

End Sub


'@Description("Returns the seconds a Peek or Poke method takes for a number of calls.")
Private Function PeekPokeTime(ByVal methodNo As PeekPokeMethod, ByVal callCount As Long) As Double
' One loop per method: a Select Case inside the loop would be timed too.
' PokeObj writes to an object variable; the strings have their own table.

    Dim SlotBool As Boolean: SlotBool = True
    Dim BoolAddress As LongPtr: BoolAddress = VarPtr(SlotBool)
    Dim ResultBool As Boolean
    Dim SlotByte As Byte: SlotByte = 123
    Dim ByteAddress As LongPtr: ByteAddress = VarPtr(SlotByte)
    Dim ResultByte As Byte
    Dim SlotCur As Currency: SlotCur = 123.4567@
    Dim CurAddress As LongPtr: CurAddress = VarPtr(SlotCur)
    Dim ResultCur As Currency
    Dim SlotDbl As Double: SlotDbl = 1.5
    Dim DblAddress As LongPtr: DblAddress = VarPtr(SlotDbl)
    Dim ResultDbl As Double
    Dim SlotInt As Integer: SlotInt = 12345
    Dim IntAddress As LongPtr: IntAddress = VarPtr(SlotInt)
    Dim ResultInt As Integer
    Dim SlotLng As Long: SlotLng = 123456
    Dim LngAddress As LongPtr: LngAddress = VarPtr(SlotLng)
    Dim ResultLng As Long
#If Win64 Then
    Dim SlotLngLng As LongLong: SlotLngLng = 123456
    Dim LngLngAddress As LongPtr: LngLngAddress = VarPtr(SlotLngLng)
    Dim ResultLngLng As LongLong
#End If
    Dim SlotPtr As LongPtr: SlotPtr = 123456
    Dim PtrAddress As LongPtr: PtrAddress = VarPtr(SlotPtr)
    Dim ResultPtr As LongPtr
    Dim SlotSng As Single: SlotSng = 1.5!
    Dim SngAddress As LongPtr: SngAddress = VarPtr(SlotSng)
    Dim ResultSng As Single
    Dim SlotVar As Variant: SlotVar = 123456
    Dim VarAddress As LongPtr: VarAddress = VarPtr(SlotVar)
    Dim ResultVar As Variant: ResultVar = 654321
    Dim SlotObj As Object: Set SlotObj = New Collection
    Dim ObjAddress As LongPtr: ObjAddress = ObjPtr(SlotObj)
    '@Ignore ImplicitUnboundDefaultMemberAccess
    Dim ObjRefAddress As LongPtr: ObjRefAddress = VarPtr(SlotObj)
    Dim ResultObj As Object: Set ResultObj = SlotObj
    Dim Pass As Long

    ' One call to initialize DMA, so that it is not timed.
    ResultLng = PeekLng(LngAddress)

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case pmBaseline
        '@Ignore EmptyForLoopBlock
        For Pass = 1 To callCount
        Next
    Case pmPeekBool
        For Pass = 1 To callCount
            ResultBool = PeekBool(BoolAddress)
        Next
    Case pmApiPeekBool
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultBool), BoolAddress, 2
        Next
    Case pmPokeBool
        For Pass = 1 To callCount
            PokeBool BoolAddress, ResultBool
        Next
    Case pmApiPokeBool
        For Pass = 1 To callCount
            ApiCopy BoolAddress, VarPtr(ResultBool), 2
        Next
    Case pmPeekByte
        For Pass = 1 To callCount
            ResultByte = PeekByte(ByteAddress)
        Next
    Case pmApiPeekByte
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultByte), ByteAddress, 1
        Next
    Case pmPokeByte
        For Pass = 1 To callCount
            PokeByte ByteAddress, ResultByte
        Next
    Case pmApiPokeByte
        For Pass = 1 To callCount
            ApiCopy ByteAddress, VarPtr(ResultByte), 1
        Next
    Case pmPeekCur
        For Pass = 1 To callCount
            ResultCur = PeekCur(CurAddress)
        Next
    Case pmApiPeekCur
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultCur), CurAddress, 8
        Next
    Case pmPokeCur
        For Pass = 1 To callCount
            PokeCur CurAddress, ResultCur
        Next
    Case pmApiPokeCur
        For Pass = 1 To callCount
            ApiCopy CurAddress, VarPtr(ResultCur), 8
        Next
    Case pmPeekDbl
        For Pass = 1 To callCount
            ResultDbl = PeekDbl(DblAddress)
        Next
    Case pmApiPeekDbl
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultDbl), DblAddress, 8
        Next
    Case pmPokeDbl
        For Pass = 1 To callCount
            PokeDbl DblAddress, ResultDbl
        Next
    Case pmApiPokeDbl
        For Pass = 1 To callCount
            ApiCopy DblAddress, VarPtr(ResultDbl), 8
        Next
    Case pmPeekInt
        For Pass = 1 To callCount
            ResultInt = PeekInt(IntAddress)
        Next
    Case pmApiPeekInt
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultInt), IntAddress, 2
        Next
    Case pmPokeInt
        For Pass = 1 To callCount
            PokeInt IntAddress, ResultInt
        Next
    Case pmApiPokeInt
        For Pass = 1 To callCount
            ApiCopy IntAddress, VarPtr(ResultInt), 2
        Next
    Case pmPeekLng
        For Pass = 1 To callCount
            ResultLng = PeekLng(LngAddress)
        Next
    Case pmApiPeekLng
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultLng), LngAddress, 4
        Next
    Case pmPokeLng
        For Pass = 1 To callCount
            PokeLng LngAddress, ResultLng
        Next
    Case pmApiPokeLng
        For Pass = 1 To callCount
            ApiCopy LngAddress, VarPtr(ResultLng), 4
        Next
#If Win64 Then
    Case pmPeekLngLng
        For Pass = 1 To callCount
            ResultLngLng = PeekLngLng(LngLngAddress)
        Next
    Case pmApiPeekLngLng
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultLngLng), LngLngAddress, 8
        Next
    Case pmPokeLngLng
        For Pass = 1 To callCount
            PokeLngLng LngLngAddress, ResultLngLng
        Next
    Case pmApiPokeLngLng
        For Pass = 1 To callCount
            ApiCopy LngLngAddress, VarPtr(ResultLngLng), 8
        Next
#End If
    Case pmPeekPtr
        For Pass = 1 To callCount
            ResultPtr = PeekPtr(PtrAddress)
        Next
    Case pmApiPeekPtr
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultPtr), PtrAddress, PtrBytes
        Next
    Case pmPokePtr
        For Pass = 1 To callCount
            PokePtr PtrAddress, ResultPtr
        Next
    Case pmApiPokePtr
        For Pass = 1 To callCount
            ApiCopy PtrAddress, VarPtr(ResultPtr), PtrBytes
        Next
    Case pmPeekSng
        For Pass = 1 To callCount
            ResultSng = PeekSng(SngAddress)
        Next
    Case pmApiPeekSng
        For Pass = 1 To callCount
            ApiCopy VarPtr(ResultSng), SngAddress, 4
        Next
    Case pmPokeSng
        For Pass = 1 To callCount
            PokeSng SngAddress, ResultSng
        Next
    Case pmApiPokeSng
        For Pass = 1 To callCount
            ApiCopy SngAddress, VarPtr(ResultSng), 4
        Next
    Case pmPeekVar
        For Pass = 1 To callCount
            ResultVar = PeekVar(VarAddress)
        Next
    Case pmPokeVar
        For Pass = 1 To callCount
            PokeVar VarAddress, ResultVar
        Next
    Case pmPeekObj
        For Pass = 1 To callCount
            Set ResultObj = PeekObj(ObjAddress)
        Next
    Case pmPokeObj
        For Pass = 1 To callCount
            PokeObj ObjRefAddress, ResultObj
        Next
    End Select
    PeekPokeTime = Stopwatch.Halt

End Function

'@Description("Returns the seconds a string method takes for a number of calls.")
Private Function StringTime( _
    ByVal methodNo As StringMethod, _
    ByVal charCount As Long, _
    ByVal callCount As Long _
) As Double
' PokeStr overwrites the target BSTR at its own length, so both strings are as long.

    Dim SlotStr As String: SlotStr = String$(charCount, "a")
    Dim StrAddress As LongPtr: StrAddress = StrPtr(SlotStr)
    Dim ResultStr As String: ResultStr = String$(charCount, "b")
    Dim Pass As Long

    ' One call to initialize DMA, so that it is not timed.
    ResultStr = PeekStr(StrAddress)

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case tmBaseline
        '@Ignore EmptyForLoopBlock
        For Pass = 1 To callCount
        Next
    Case tmPeekStr
        For Pass = 1 To callCount
            ResultStr = PeekStr(StrAddress)
        Next
    Case tmPokeStr
        For Pass = 1 To callCount
            PokeStr StrAddress, ResultStr
        Next
    Case tmAssign
        For Pass = 1 To callCount
            ResultStr = SlotStr
        Next
    End Select
    StringTime = Stopwatch.Halt

End Function


'@Description("Returns the seconds an array method takes for a number of calls.")
Private Function ArrayTime( _
    ByVal methodNo As ArrayMethod, _
    ByRef subject As Variant, _
    ByVal callCount As Long _
) As Double

    Dim ResultPtr As LongPtr
    Dim ResultDims As Integer
    Dim ResultFlag As Boolean
    Dim Pass As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case amBaseline
        '@Ignore EmptyForLoopBlock
        For Pass = 1 To callCount
        Next
    Case amArrPtr
        For Pass = 1 To callCount
            ResultPtr = ArrPtr(subject)
        Next
    Case amArrVarPtr
        For Pass = 1 To callCount
            ResultPtr = ArrVarPtr(subject)
        Next
    Case amArrDataPtr
        For Pass = 1 To callCount
            ResultPtr = ArrDataPtr(subject)
        Next
    Case amArrDims
        For Pass = 1 To callCount
            ResultDims = ArrDims(subject)
        Next
    Case amIsArrayEmpty
        For Pass = 1 To callCount
            ResultFlag = IsArrayEmpty(subject)
        Next
    End Select
    ArrayTime = Stopwatch.Halt

End Function


'@Description("Returns the seconds a slice method takes for a number of calls.")
Private Function SliceTime( _
    ByVal methodNo As SliceMethod, _
    ByVal elementCount As Long, _
    ByVal callCount As Long _
) As Double
' The arrays are local: passing a ByRef array parameter on to a Property Let
' can crash VBA.

    Dim FromValues() As Long
    ReDim FromValues(1 To elementCount)
    Dim ToValues() As Long
    ReDim ToValues(1 To elementCount)
    Dim Position As Long
    For Position = 1 To elementCount
        FromValues(Position) = Position
    Next
    Dim Slice As Variant
    Dim Pass As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case smGetMidArr
        For Pass = 1 To callCount
            Slice = MidArr(FromValues, 1)
        Next
    Case smGetLoop
        For Pass = 1 To callCount
            ReDim ToValues(1 To elementCount)
            For Position = 1 To elementCount
                ToValues(Position) = FromValues(Position)
            Next
        Next
    Case smLetMidArr
        For Pass = 1 To callCount
            MidArr(ToValues, 1) = FromValues
        Next
    Case smLetLoop
        For Pass = 1 To callCount
            For Position = 1 To elementCount
                ToValues(Position) = FromValues(Position)
            Next
        Next
    End Select
    SliceTime = Stopwatch.Halt

End Function


'@Description("Returns the seconds a memory block method takes for a number of calls.")
Private Function BlockTime( _
    ByVal methodNo As BlockMethod, _
    ByVal byteCount As Long, _
    ByVal callCount As Long _
) As Double

    Dim FromValues() As Byte
    ReDim FromValues(0 To byteCount - 1)
    Dim ToValues() As Byte
    ReDim ToValues(0 To byteCount - 1)
    Dim FromAt As LongPtr: FromAt = VarPtr(FromValues(0))
    Dim ToAt As LongPtr: ToAt = VarPtr(ToValues(0))
    Dim Pass As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case bmCopyMem
        For Pass = 1 To callCount
            CopyMem ToAt, FromAt, byteCount
        Next
    Case bmLSetCopy
        For Pass = 1 To callCount
            LSetCopyMem ToAt, FromAt, byteCount
        Next
    Case bmApiCopy
        For Pass = 1 To callCount
            ApiCopy ToAt, FromAt, byteCount
        Next
    Case bmZeroMem
        For Pass = 1 To callCount
            ZeroMem ToAt, byteCount
        Next
    Case bmLSetZero
        For Pass = 1 To callCount
            LSetZeroMem ToAt, byteCount
        Next
    Case bmApiZero
        For Pass = 1 To callCount
            ApiZero ToAt, byteCount
        Next
    End Select
    BlockTime = Stopwatch.Halt

End Function

'@Description("Copies a memory block like CopyMem, with the LSet method for every size.")
Private Sub LSetCopyMem( _
    ByVal targetAddress As LongPtr, _
    ByVal sourceAddress As LongPtr, _
    ByVal byteCount As Long _
)
' CopyMem without the switch to CopyMemory: the two blocks pose as BSTR strings
' (length prefix in their first 4 bytes) and LSet copies them, which keeps the
' target's length. No overlap check; no error may occur while the strings point
' into the blocks, so they are reset to vbNullString straight after the LSet.

    If byteCount <= 0 Then Exit Sub

    Dim EvenCount As Long: EvenCount = byteCount
    ' Copy the last byte if there is an uneven number of bytes.
    If EvenCount Mod 2 = 1 Then
        EvenCount = EvenCount - 1
        PokeByte targetAddress + EvenCount, PeekByte(sourceAddress + EvenCount)
    End If
    ' Copy the remaining even number of bytes.
    Select Case EvenCount
    Case Is >= 10
        Dim FirstLong As Long: FirstLong = PeekLng(sourceAddress)
        PokeLng sourceAddress, EvenCount - 4
        PokeLng targetAddress, EvenCount - 4
        Dim TargetText As String
        Dim SourceText As String
        PokePtr VarPtr(TargetText), targetAddress + 4
        PokePtr VarPtr(SourceText), sourceAddress + 4
        LSet TargetText = SourceText
        PokePtr VarPtr(TargetText), 0
        PokePtr VarPtr(SourceText), 0
        PokeLng sourceAddress, FirstLong
        PokeLng targetAddress, FirstLong
    Case 8
        PokeCur targetAddress, PeekCur(sourceAddress)
    Case 6
        PokeLng targetAddress, PeekLng(sourceAddress)
        PokeInt targetAddress + 4, PeekInt(sourceAddress + 4)
    Case 4
        PokeLng targetAddress, PeekLng(sourceAddress)
    Case 2
        PokeInt targetAddress, PeekInt(sourceAddress)
    End Select

End Sub


'@Description("Zeroes a memory block like ZeroMem, with the LSet method for every size.")
Private Sub LSetZeroMem(ByVal targetAddress As LongPtr, ByVal byteCount As Long)
' ZeroMem without the switch to ZeroMemory: the block poses as a BSTR string and
' LSet copies a string of null characters into it.

    If byteCount <= 0 Then Exit Sub

    Dim EvenCount As Long: EvenCount = byteCount
    ' Zero the last byte if there is an uneven number of bytes.
    If EvenCount Mod 2 = 1 Then
        EvenCount = EvenCount - 1
        PokeByte targetAddress + EvenCount, 0
    End If
    ' Zero the remaining even number of bytes.
    Select Case EvenCount
    Case Is >= 10
        PokeLng targetAddress, EvenCount - 4
        Dim TargetText As String
        PokePtr VarPtr(TargetText), targetAddress + 4
        LSet TargetText = String$((EvenCount - 4) \ 2, vbNullChar)
        PokePtr VarPtr(TargetText), 0
        PokeLng targetAddress, 0
    Case 8
        PokeCur targetAddress, 0
    Case 6
        PokeLng targetAddress, 0
        PokeInt targetAddress + 4, 0
    Case 4
        PokeLng targetAddress, 0
    Case 2
        PokeInt targetAddress, 0
    End Select

End Sub


'@Description("Returns the seconds a collection method takes for a number of calls.")
Private Function CollectionTime( _
    ByVal methodNo As CollectionMethod, _
    ByVal keyedItems As Collection, _
    ByVal callCount As Long _
) As Double

    Dim Middle As Long: Middle = keyedItems.Count \ 2
    Dim MiddleKey As String: MiddleKey = "K" & Middle
    Dim ResultVar As Variant
    Dim ResultStr As String
    Dim ResultPos As Long
    Dim ResultKeys() As String
    Dim Pass As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case cmItemByPosition
        For Pass = 1 To callCount
            ResultVar = keyedItems.Item(Middle)
        Next
    Case cmItemByKey
        For Pass = 1 To callCount
            ResultVar = keyedItems.Item(MiddleKey)
        Next
    Case cmCollKey
        For Pass = 1 To callCount
            ResultStr = CollKey(keyedItems, Middle)
        Next
    Case cmCollPos
        For Pass = 1 To callCount
            ResultPos = CollPos(keyedItems, MiddleKey)
        Next
    Case cmCollKeys
        For Pass = 1 To callCount
            ResultKeys = CollKeys(keyedItems)
        Next
    End Select
    CollectionTime = Stopwatch.Halt

End Function
