Attribute VB_Name = "DMA"
Attribute VB_Description = "Direct Memory Access Module."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Direct Memory Access Module.
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.09.10
'
' This module implements Peek and Poke methods for intrinsic data types as well
' as a set of auxiliary methods for arrays, collections and memory blocks.
'
' This module augments the standard VB Collection class which misses some useful
' methods like getting the keys for the items contained in a Collection. The is
' no VB function available to get the keys other than my poking around memory.
'
' This module provides a series of useful methods for VB Arrays which are not
' standard available, e.g. a function to obtain the number of dimensions, a
' function to check whether an array is empty, a function to get the array
' pointer, and a function to create an array of any dimensions at run time, etc.
'
' Using API's for memory access can be very slow depending on specific system
' settings. Christian Buse came up with the idea to use a SafeArray construct
' for memory access, which is faster than the more common Variant ByRef method,
' mainly because the latter requires at least one extra stack frame for each
' memory access. see: https://github.com/cristianbuse/VBA-MemoryTools.
' Contrary to Christian Buse's LibMemory, the PeekPoke uses one construct common
' to all methods whilst taking extra measurements to allow for nested calls.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methods for peeking and poking intrinsic variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Methods
' PeekBool(address)             Returns the Boolean located at a memory address
' PeekByte(address)             Returns the Byte located at a memory address
' PeekCur(address)              Returns the Currency located at a memory address
' PeekDbl(address)              Returns the Double located at a memory address
' PeekInt(address)              Returns the Integer located at a memory address
' PeekLng(address)              Returns the Long located at a memory address
' PeekLngLng(address)(x64)      Returns the LongLong located at a memory address
' PeekPtr(address)              Returns the LongPtr located at a memory address
' PeekSng(address)              Returns the Single located at a memory address
' PeekVar(address)              Returns the Variant located at a memory address
' PeekStr(address)              Returns the String located at a memory address
' PeekObj(address)              Returns the Object located at a memory address
' PokeBool(address, value)      Copies a Boolean to a memory address
' PokeCur(address, value)       Copies a Currency to a memory address
' PokeByte(address, value)      Copies a Byte to a memory address
' PokeDbl(address, value)       Copies a Double to a memory address
' PokeInt(address, value)       Copies a Integer to a memory address
' PokeLng(address, value)       Copies a Long to a memory address
' PokeLngLng(address, value)    Copies a LongLong to a memory address
' PokePtr(address, value)       Copies a LongPtr to a memory address
' PokeSng(address, value)       Copies a Single to a memory address
' PokeVar(address, value)       Copies a Variant to a memory address
' PokeStr(address, value)       Copies a String to a memory address
' PokeObj(address, value)       Copies an Object to a memory address
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methods and properties for arrays
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Methods
' ArrPtr(arr)                   Returns the address of the underlying SafeArray
' ArrVarPtr(arr)                Returns the address of an array variable
' ArrDataPtr(arr)               Returns the address of the stored data
' CreateArray(vt [, dims])      Returns a new array of specified dimensions
' CloneArray(address)           Returns an array clone of a SafeArray at an address
'
' Properties (Get)
' ArrDims(arr)                  Returns the number of dimensions of an array
' ArrElemSize(arr)              Returns the size of an element of an array
' IsArrayAllocated(arr)         Returns True if an array is allocated or False if not
' IsArrayEmpty(arr)             Returns True if an array is empty or False if not
' IsArrayStatic(arr)            Returns True if an array is static or False if not
'
' Properties (Let)
' ArrBase(arr [, dim])          Sets the lower boundary of an array dimension
'
' Properties (Set/Let)
' ArrElem(arr, indices)         Gets or sets an array element by index vector
' MidArr(arr, start [, length]) Gets or sets an array slice
'
' Timings (ms) for array procedures.
' Method      Long(100)      Long() Variant(100)  Variant()
' ArrPtr        0.00009     0.00009     0.00009     0.00009
' ArrVarPtr     0.00008     0.00008     0.00008     0.00008
' ArrDataPtr    0.00018     0.00014     0.00013     0.00015
' ArrDims       0.00022     0.00013     0.00014     0.00013
'
' Timings (ms) for MidArr for Long(1 to length) array.
'    Length  Get MidArr    For loop  Let MidArr    For loop
'         0     0.00056     0.00363     0.00226     0.00445
'        10     0.00057     0.00372     0.00267     0.00475
'       100     0.00057     0.00624     0.00263     0.00858
'     1.000     0.00061     0.03157     0.00282     0.04593
'    10.000     0.00160     0.30272     0.00494     0.43832
'   100.000     0.01241     2.97955     0.04005     4.33674
'
' MidArr as well as a For loop take significant longer for Strings and Variants.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methods for memory blocks
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Methods
' CopyMem destination, source, NBytes      Copies a memory block
' ZeroMem address, NBytes       Copies null bytes to a memory block
' DumpMem(address, NBytes)      Returns a hexadecimal dump of a memory block
'
' CopyMem and ZeroMem can be used to substitute CopyMemory and ZeroMemory API's.
' Some code adaptation may be required because their arguments always need to
' be provided as pointers, i.e. the use of Any is only valid for an APIs.
'
' Timings (ms) for comparison of the CopyMem Method and CopyMemory API.
'        NBytes             CopyMem  CopyMemory
'            10             0.00035     0.00152
'           100             0.00032     0.00155
'         1.000             0.00037     0.00161
'        10.000             0.00053     0.00169
'       100.000             0.00275     0.00302
'     1.000.000             0.06700     0.06945
'    10.000.000             0.70597     0.68546 <- break even point
'
' Timings (ms) for comparison of the ZeroMem Method and ZeroMemory API.
'        NBytes             ZeroMem  ZeroMemory
'            10             0.00021     0.00185
'           100             0.00021     0.00163
'         1.000             0.00034     0.00162
'        10.000             0.00153     0.00172 <- break even point
'       100.000             0.01506     0.00152
'     1.000.000             0.19999     0.01422
'    10.000.000            73.39400     0.73259
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methods and properties for VB Collections
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Methods
' CollKeys(coll [, base])       Returns an array containing all the keys
' CollKey(coll, pos)            Returns the key of coll.item(pos)
' CollPos(coll, key)            Returns the position of coll.item(key)
'
' Properties (Let)
' CollItem(coll, index)         Assigns a new value to coll.item(index)
'
' The standard VB Collection class misses some useful methods, e.g. see:
' https://stackoverflow.com/questions/5702362/vba-collection-list-of-keys
'
' Direct memory access of the double linked list data structure of a Collection
' makes it possible to get access to the otherwise hidden keys. Access to the
' keys is essential to serialize a Collection or to sort a Collection by key.
' Although a Dictionary is a good alternative for a Collection in many ways, it
' actually misses the flexibility to handle mixed items with and without keys.
' The Dictionary Class implements a rather primitive enumeration via an array of
' keys or items. The methods provided here neatly augment a VB Collection Class.
'
' All these methods are inherently "slow" as they require to traverse the double
' linked list and in fact have similar performance challenges as for looking up
' a Collection element by its position. I did not manage to crack the hash table
' in order to get direct access to a node for a given key.
'
' Timings (ms) for object.Item method and the Collection extensions.
'    Count   .Item i  .Item key        CollKey   CollPos  CollKeys
'       10    0.0001     0.0002         0.0004    0.0011    0.0011
'      100    0.0001     0.0002         0.0006    0.0049    0.0094
'    1.000    0.0011     0.0005         0.0042    0.0449    0.0847
'   10.000    0.0128     0.0008         0.0360    0.4129    0.8163
'  100.000    0.1773     0.0010         0.3288    4.2147    9.8930
'
' The results are for an item at position Count\2 and all items have a key. The
' Item key method uses a hash table to look up the double linked list node and
' is thus much faster for a large collections than the Item method which has to
' traverse the list. CollItemKey is slower than the Item method because it uses
' VB code rather than internal code. CollItemPos is much slower because it not
' only has to traverse the linked list, but it also needs to compare the key
' with the keys stored in the Collection. CollKeys has to traverse the whole
' linked list and needs to allocate strings for all the Collection keys.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The null pointer constant (not VB defined).
Public Const vbNullPtr As LongPtr = 0

' The variable subtype constant for a LongPtr (not VB defined).
#If Win64 Then
Public Const vbLongPtr As Integer = VBA.vbLongLong
#Else
Public Const vbLongPtr As Integer = VBA.vbLong
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private API declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' VarPtrArray returns the pointer to an array variable.
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" ( _
    var() As Any _
) As LongPtr

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlmovememory
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    pDst As Any, _
    pSrc As Any, _
    ByVal NBytes As Long _
)

' https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlzeromemory
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" ( _
    pDst As Any, _
    ByVal NBytes As Long _
)

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraycreate
Private Declare PtrSafe Function SafeArrayCreate Lib "oleaut32.dll" ( _
    ByVal vt As Integer, _
    ByVal cDims As Integer, _
    ByVal rgsabound As LongPtr _
) As LongPtr

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearraygetelement
Private Declare PtrSafe Function SafeArrayGetElement Lib "oleaut32.dll" ( _
    ByVal psa As LongPtr, _
    ByVal rgIndices As LongPtr, _
    ByRef pv As Any _
) As Long

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-safearrayputelement
Private Declare PtrSafe Function SafeArrayPutElement Lib "oleaut32.dll" ( _
    ByVal psa As LongPtr, _
    ByVal rgIndices As LongPtr, _
    ByRef pv As Any _
) As Long

' https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-variantchangetype
Private Declare PtrSafe Function VariantChangeType Lib "oleaut32.dll" ( _
    ByRef pvDst As Variant, _
    ByRef pvSrc As Variant, _
    ByVal wFlags As Integer, _
    ByVal vt As Integer _
) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The memory size of intrinsic data types.
' https://docs.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/data-type-summary
Private Enum VARSIZE
    vbSizeByte = 1
    vbSizeBoolean = 2
    vbSizeInteger = 2
    vbSizeLong = 4
    vbSizeSingle = 4
    vbSizeDouble = 8
    vbSizeCurrency = 8
    vbSizeDate = 8
#If Win64 Then
    vbSizeLongLong = 8
    vbSizeVariant = 24
    vbSizeLongPtr = vbSizeLongLong
#Else
    vbSizeVariant = 16
    vbSizeLongPtr = vbSizeLong
#End If
End Enum

' The memory offsets for intrinsic data types (implementation dependent).
Private Enum MEMOFFSET
    ' The memory offsets from the VarPtr(var) address of a Variant.
    vbOffsetVariantType = 0             ' type (Integer)
    vbOffsetVariantDecimal = 2          ' data (Decimal only - 14 bits)
    vbOffsetVariantData = 8             ' *data (LongPtr) or data
    ' The memory offset from the StrPtr(var) address of a BSTR.
    vbOffsetStringLenB = -4             ' number of bytes (Long) (= LenB)
#If Win64 Then
    ' The memory offsets from the ObjPtr(var) address of a Collection.
    vbOffsetCollectionvTable = 0        ' *vTable (see COM)
    vbOffsetCollectionnRef = 24         ' nref (Long)
    vbOffsetCollectionCount = 28        ' count (Long)
    vbOffsetCollectionHead = 40         ' *head (LongPtr)
    vbOffsetCollectionTail = 48         ' *tail (LongPtr)
    ' The memory offsets from the node address in a Collection object.
    vbOffsetCollectionItem = 0          ' item (Variant)
    vbOffsetCollectionKey = 24          ' *key  (LongPtr)
    vbOffsetCollectionPrev = 32         ' *prev (LongPtr)
    vbOffsetCollectionNext = 40         ' *next (LongPtr)
    ' The memory offsets from the ArrPtr(var) address of a SafeArray
    vbOffsetArrayType = -4              ' VarType (Long) for FADF_HAVEVARTYPE = 1
    vbOffsetArraycDims = 0              ' number of dimensions (Integer)
    vbOffsetArrayfFeatures = 2          ' flags (Integer)
    vbOffsetArraycbElements = 4         ' element size (Long)
    vbOffsetArraycLocks = 8             ' lock count (Long)
    ' For x64 bits 12-15 are not used.
    vbOffsetArraypvData = 16            ' *data (LongPtr)
    vbOffsetArrayrgsaBound = 24         ' cElements, lLbound pairs (Long()).
#Else
    vbOffsetCollectionvTable = 0
    vbOffsetCollectionnRef = 12
    vbOffsetCollectionCount = 16
    vbOffsetCollectionHead = 24
    vbOffsetCollectionTail = 28
    vbOffsetCollectionItem = 0
    vbOffsetCollectionKey = 16
    vbOffsetCollectionPrev = 20
    vbOffsetCollectionNext = 24
    vbOffsetArrayType = -4
    vbOffsetArraycDims = 0
    vbOffsetArrayfFeatures = 2
    vbOffsetArraycbElements = 4
    vbOffsetArraycLocks = 8
    vbOffsetArraypvData = 12
    vbOffsetArrayrgsaBound = 16
#End If
End Enum

' https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum
Private Enum VARENUM
    VT_BYREF = &H4000
    VT_TYPEMASK = &HFFF
End Enum

' https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray
Private Enum ADVFEATUREFLAGS
    FADF_AUTO = &H1
    FADF_STATIC = &H2
    FADF_EMBEDDED = &H4
    FADF_FIXEDSIZE = &H10
    FADF_RECORD = &H20
    FADF_HAVEIID = &H40
    FADF_HAVEVARTYPE = &H80
    FADF_BSTR = &H100
    FADF_UNKNOWN = &H200
    FADF_DISPATCH = &H400
    FADF_VARIANT = &H800
    FADF_RESERVED = &HF008
End Enum

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorSubscriptOutOfRange = 9
    vbErrorTypeMismatch = 13
    vbErrorInternalError = 51
    vbErrorObjectVariableNotSet = 91
    vbErrorObjectRequired = 424
    vbErrorArgumentNotOptional = 449
End Enum

' https://learn.microsoft.com/en-us/windows/win32/api/oaidl/ns-oaidl-safearray
' The SAFEARRAY UDT is a combination of a SAFEARRAY and a SAFEARRAYBOUND structure.
Private Type SAFEARRAY
    cDims As Integer            ' The number of dimensions.
    fFeatures As Integer        ' Advanced features.
    cbElements As Long          ' The size of an element.
    cLocks As Long              ' Lock count.
    pvData As LongPtr           ' Pointer to the data.
    cElements As Long           ' Number of elements.
    lLbound As Long             ' Lower bound of index.
End Type

' A special construct is used for memory access by pointer. Initially, all type
' fields except the last are non-allocated arrays. Initialization points all the
' arrays to the last MetaData (SAFEARRAY) field. For peeking or poking the data
' at a memory location, the MetaData data pointer is temporarily directed to the
' memory address. The construct is put in an idle state by setting the array
' data pointer to an empty array. This avoids triggering the automatic garbage
' collection process which would cause a system crash.

Private Type CONSTRUCT
    Bytes() As Byte
    Booleans() As Boolean
    Currencies() As Currency
    Doubles() As Double
    Integers() As Integer
    Longs() As Long
    Objects() As Object
    Pointers() As LongPtr
    Singles() As Single
    Strings() As String
    Variants() As Variant
    SafeArrays() As SAFEARRAY   ' Used to copy a SAFEARRAY at once.
    MetaData As SAFEARRAY
End Type
Private this As CONSTRUCT


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Comments on  using a COMSTRUCT.
' Saving and restoring the previous values for this.MetaData.cElements and
' this.MetaData.pvData allows nested call of memory manipulating procedures,
' because they all share one common CONSTRUCT. Alternatively each procedure
' would have to store an individual CONSTRUCT (like in LibMemory).

Public Function OffsetPtr( _
    ByVal address As LongPtr, _
    ByVal Offset As Long _
) As LongPtr
Attribute OffsetPtr.VB_Description = "Returns the offset address of a memory address."
' Safe pointer arithmetic for signed numbers (not required normally).

    Const SignBitMask As LongPtr = -(2 ^ (8 * vbSizeLongPtr - 1))
    OffsetPtr = (address Xor SignBitMask) + Offset Xor SignBitMask

End Function


Public Function PeekBool(ByVal address As LongPtr) As Boolean
Attribute PeekBool.VB_Description = "Returns the Boolean located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekBool = this.Booleans(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeBool(ByVal address As LongPtr, ByVal value As Boolean)
Attribute PokeBool.VB_Description = "Copies a Boolean to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Booleans(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekByte(ByVal address As LongPtr) As Byte
Attribute PeekByte.VB_Description = "Returns the Byte located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekByte = this.Bytes(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeByte(ByVal address As LongPtr, ByVal value As Byte)
Attribute PokeByte.VB_Description = "Copies a Byte to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Bytes(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekCur(ByVal address As LongPtr) As Currency
Attribute PeekCur.VB_Description = "Returns the Currency located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekCur = this.Currencies(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeCur(ByVal address As LongPtr, ByVal value As Currency)
Attribute PokeCur.VB_Description = "Copies a Currency to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Currencies(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekDbl(ByVal address As LongPtr) As Double
Attribute PeekDbl.VB_Description = "Returns the Double located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekDbl = this.Doubles(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeDbl(ByVal address As LongPtr, ByVal value As Double)
Attribute PokeDbl.VB_Description = "Copies a Double to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Doubles(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekInt(ByVal address As LongPtr) As Integer
Attribute PeekInt.VB_Description = "Returns the Integer located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekInt = this.Integers(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeInt(ByVal address As LongPtr, ByVal value As Integer)
Attribute PokeInt.VB_Description = "Copies an Integer to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Integers(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekLng(ByVal address As LongPtr) As Long
Attribute PeekLng.VB_Description = "Returns the Long located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekLng = this.Longs(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeLng(ByVal address As LongPtr, ByVal value As Long)
Attribute PokeLng.VB_Description = "Copies a Long to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Longs(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


#If Win64 Then
Public Function PeekLngLng(ByVal address As LongPtr) As LongLong
Attribute PeekLngLng.VB_Description = "Returns the LongLong located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekLngLng = this.Pointers(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function
#End If


#If Win64 Then
Public Sub PokeLngLng(ByVal address As LongPtr, ByVal value As LongLong)
Attribute PokeLngLng.VB_Description = "Copies a LongLong to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Pointers(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub
#End If


Public Function PeekPtr(ByVal address As LongPtr) As LongPtr
Attribute PeekPtr.VB_Description = "Returns the LongPtr located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekPtr = this.Pointers(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokePtr(ByVal address As LongPtr, ByVal value As LongPtr)
Attribute PokePtr.VB_Description = "Copies a LongPtr to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Pointers(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekSng(ByVal address As LongPtr) As Single
Attribute PeekSng.VB_Description = "Returns the Single located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekSng = this.Singles(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeSng(ByVal address As LongPtr, ByVal value As Single)
Attribute PokeSng.VB_Description = "Copies a Single to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.Singles(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekVar(ByVal address As LongPtr) As Variant
Attribute PeekVar.VB_Description = "Returns the Variant located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    If VBA.IsObject(this.Variants(0)) Then
        Set PeekVar = this.Variants(0)
    Else
        PeekVar = this.Variants(0)
    End If
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeVar(ByVal address As LongPtr, ByRef value As Variant)
Attribute PokeVar.VB_Description = "Copies a Variant to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    If VBA.IsObject(value) Then
        Set this.Variants(0) = value
    Else
        this.Variants(0) = value
    End If
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekStr(ByVal address As LongPtr) As String
Attribute PeekStr.VB_Description = "Returns the String located at a memory address."
' Address is a BSTR address (StrPtr).

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(address)
    PeekStr = this.Strings(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeStr(ByVal address As LongPtr, ByVal value As String)
Attribute PokeStr.VB_Description = "Copies a String to a memory address."
' Address is a BSTR address (StrPtr).

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(address)
    LSet this.Strings(0) = value ' Keeps original size of BSTR string.
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function PeekObj(ByVal address As LongPtr) As Object
Attribute PeekObj.VB_Description = "Returns the Object located at a memory address."
' Address is an object address (ObjPtr).

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(address)
    Set PeekObj = this.Objects(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Sub PokeObj(ByVal address As LongPtr, ByVal value As Object)
Attribute PokeObj.VB_Description = "Copies a Object to a memory address."
' Address is the address of the object reference.

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    Set this.Objects(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function ArrVarPtr(ByRef arr As Variant) As LongPtr
Attribute ArrVarPtr.VB_Description = "Returns the address of an array variable."
' Unlike the VarPtrArray API, ArrVarPtr accepts a variant/array as argument.

    If VBA.IsArray(arr) = False Then Err.Raise vbErrorTypeMismatch

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(arr)
    If (this.Integers(0) And VT_BYREF) = VT_BYREF Then
        ' The argument provided is an array.
        this.MetaData.pvData = this.MetaData.pvData + vbOffsetVariantData
        ArrVarPtr = this.Pointers(0)
    Else
        ' The argument provided is a variant/array.
        ArrVarPtr = this.MetaData.pvData + vbOffsetVariantData
    End If
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Function ArrPtr(ByRef arr As Variant) As LongPtr
Attribute ArrPtr.VB_Description = "Returns the address of the SafeArray of an array."
' Returns vbNullPtr for an unallocated array.

    If VBA.IsArray(arr) = False Then Err.Raise vbErrorTypeMismatch

    If this.MetaData.cDims = 0 Then InitializeConstruct

    ' Repeat the ArrVarPtr code to avoid a stack frame.
    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(arr)
    If (this.Integers(0) And VT_BYREF) = VT_BYREF Then
        this.MetaData.pvData = this.MetaData.pvData + vbOffsetVariantData
        this.MetaData.pvData = this.Pointers(0)
    Else
        this.MetaData.pvData = this.MetaData.pvData + vbOffsetVariantData
    End If
    ArrPtr = this.Pointers(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Function ArrDataPtr(ByRef arr As Variant) As LongPtr
Attribute ArrDataPtr.VB_Description = "Returns the address of the stored data of an array."
' Returns vbNullPtr for an unallocated array.

    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa <> vbNullPtr Then
        ArrDataPtr = PeekPtr(psa + vbOffsetArraypvData)
    End If

End Function


Public Property Get ArrDims(ByRef arr As Variant) As Integer
Attribute ArrDims.VB_Description = "Returns the number of dimensions of an array."
' Returns 0 for an unallocated array.

    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa <> vbNullPtr Then
        ArrDims = PeekInt(psa + vbOffsetArraycDims)
    End If

End Property


Public Property Get ArrElemSize(ByRef arr As Variant) As Long
Attribute ArrElemSize.VB_Description = "Returns the memory size of an element of an array."
' Returns 0 for a unallocated array.

    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa <> vbNullPtr Then
        ArrElemSize = PeekLng(psa + vbOffsetArraycbElements)
    End If

End Property


Public Property Get IsArrayStatic(ByRef arr As Variant) As Boolean
Attribute IsArrayStatic.VB_Description = "Returns True if an array is static or False otherwise."
' Returns False for a unallocated array.

    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa <> vbNullPtr Then
        IsArrayStatic = (PeekInt(psa + vbOffsetArrayfFeatures) And FADF_STATIC) = FADF_STATIC
    End If

End Property


Public Property Get IsArrayAllocated(ByRef arr As Variant) As Boolean
Attribute IsArrayAllocated.VB_Description = "Returns True if an array is allocated or False otherwise."
' An array is allocated if it has an allocated SafeArray.

    Dim psa As LongPtr: psa = ArrPtr(arr)
    IsArrayAllocated = psa <> vbNullPtr

End Property


Public Property Get IsArrayEmpty(ByRef arr As Variant) As Boolean
Attribute IsArrayEmpty.VB_Description = "Returns True if an array is empty or False otherwise."
' Returns True for a unallocated array and for an allocated array without data.
' see https://www.cpearson.com/Excel/IsArrayAllocated.aspx

    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa <> vbNullPtr Then
        IsArrayEmpty = PeekLng(psa + vbOffsetArrayrgsaBound) = 0
    Else
        IsArrayEmpty = True
    End If

End Property


Public Function CreateArray( _
    ByVal vt As VBA.VbVarType, _
    ParamArray dims() As Variant _
) As Variant
Attribute CreateArray.VB_Description = "Returns a created variant/array of specified dimensions."
' CreateArray provides a solution for creating an array when its (number of)
' dimensions and type are determined during run time. See the Select Case
' below for valid options for specifying the array dimensions. CreateArray
' can also be used for creating empty typed arrays. If no optional arguments
' an unallocated array is returned, which can be allocated later with Redim.

    vt = vt And VT_TYPEMASK

    Select Case vt
    Case VBA.vbInteger To VBA.vbDataObject, VBA.vbByte, vbLongPtr
        If UBound(dims) < LBound(dims) Then
            ' Create a typed empty variant/array. No SafeArray is allocated.
            PokeInt VarPtr(CreateArray), (vt And VT_TYPEMASK) Or VBA.vbArray
            Exit Function
        End If
    Case Else
        ' vbEmpty, vbNull, vbDecimal, vbUserDefinedType or undefined types.
        Err.Raise vbErrorTypeMismatch, , "Array type is invalid"
    End Select

    Dim cDims As Integer
    If VBA.IsArray(dims(0)) Then
        ' The optional argument is expected to be a rgsbound vector.
        If UBound(dims) <> 0 Then Err.Raise vbErrorInvalidProcedureCall
        If LBound(dims(0)) <> 0 Then Err.Raise vbErrorInvalidProcedureCall
        cDims = (UBound(dims(0)) + 1) \ 2
    Else
        ' The first optional argument is expected to be NDims.
        If dims(0) <= 0 Then Err.Raise vbErrorInvalidProcedureCall
        ' followed by at least one more argument.
        If UBound(dims) = 0 Then Err.Raise vbErrorArgumentNotOptional
        cDims = dims(0)
    End If

    ' The array dimensions can be provided in multiple ways.
    Dim rgsabound() As Long: ReDim rgsabound(0 To 2 * cDims - 1)
    Dim i As Long
    Select Case UBound(dims)
    Case 0
        ' dims = rgsabound.
        For i = 0 To 2 * cDims - 1
            rgsabound(i) = dims(0)(i)
        Next
    Case 1
        ' dims = (cDims, cElements)
        For i = 0 To 2 * cDims - 1 Step 2
            rgsabound(i) = dims(1)
        Next
    Case cDims
        ' dims = (cDims, cElements(1), ... , cElements(cDims)).
        For i = 0 To 2 * cDims - 1 Step 2
            rgsabound(i) = dims(i \ 2 + 1)
        Next
    Case 2 * cDims
        ' dims = (cDims, cElements(1), lLbound(1), ... , cElements(cDims), lLbound(cDims)).
        For i = 0 To 2 * cDims - 1
            rgsabound(i) = dims(i + 1)
        Next
    Case Else
        Err.Raise vbErrorInvalidProcedureCall
    End Select

    ' Allocate the SafeArray.
    CreateArray = SafeArrayCreate(vt, cDims, VarPtr(rgsabound(0)))
    If CreateArray = vbNullPtr Then Err.Raise vbErrorInternalError

    ' Return a typed variant/array.
    PokeInt VarPtr(CreateArray), vt Or VBA.vbArray

End Function


Public Property Get ArrElem( _
    ByRef arr As Variant, _
    ByRef Indices() As Long _
) As Variant
Attribute ArrElem.VB_Description = "Returns or sets the value of an array element specified by its indices."
' ArrayElem property provides a solution to access the element of an array for with
' the dimensions are created during run time by using CreateArray.

    Dim vt As Long: vt = VBA.VarType(arr) And VT_TYPEMASK
    Select Case vt
    Case VBA.vbVariant
        If SafeArrayGetElement(ArrPtr(arr), VarPtr(Indices(0)), ArrElem) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    Case Else
        ' Copy the element to the data portion of the Variant.
        If SafeArrayGetElement(ArrPtr(arr), VarPtr(Indices(0)), ByVal VarPtr(ArrElem) + vbOffsetVariantData) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
        ' Set the type of the Variant.
        PokeInt VarPtr(ArrElem), vt
    End Select

End Property
Public Property Let ArrElem( _
    ByRef arr As Variant, _
    ByRef Indices() As Long, _
    ByVal RHS As Variant _
)

    Dim vt As Long: vt = VBA.VarType(arr) And VT_TYPEMASK
    Select Case vt
    Case VBA.vbVariant
        If SafeArrayPutElement(ArrPtr(arr), VarPtr(Indices(0)), RHS) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    Case VBA.vbString
        If VBA.VarType(RHS) <> VBA.vbString Then
            ' In situ type conversion of RHS to a string.
            RHS = VBA.CStr(RHS)
        End If
        ' Copy the RHS string to the array element.
        If SafeArrayPutElement(ArrPtr(arr), VarPtr(Indices(0)), ByVal StrPtr(RHS)) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    Case Else
        If VBA.VarType(RHS) <> vt Then
            ' In situ type conversion of RHS to the array type.
            If VariantChangeType(RHS, RHS, 0, vt) < 0 Then
                Err.Raise vbErrorTypeMismatch
            End If
        End If
        ' Copy the data portion of RHS to the array element.
        If SafeArrayPutElement(ArrPtr(arr), VarPtr(Indices(0)), ByVal VarPtr(RHS) + vbOffsetVariantData) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    End Select
End Property
Public Property Set ArrElem( _
    ByRef arr As Variant, _
    ByRef Indices() As Long, _
    ByVal RHS As Variant _
)

    Dim vt As Long: vt = VBA.VarType(arr) And VT_TYPEMASK
    Select Case vt
    Case VBA.vbVariant
        If SafeArrayPutElement(ArrPtr(arr), VarPtr(Indices(0)), RHS) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    Case VBA.vbObject
        If SafeArrayPutElement(ArrPtr(arr), VarPtr(Indices(0)), ByVal ObjPtr(RHS)) < 0 Then
            Err.Raise vbErrorSubscriptOutOfRange, , "Array indices are invalid"
        End If
    End Select

End Property


Public Property Let ArrBase( _
    ByRef arr As Variant, _
    Optional ByVal dimension As Long = 1, _
    ByVal RHS As Long _
)
Attribute ArrBase.VB_Description = "Sets the lower boundary of a dimension of an array."

    Dim cDims As Long: cDims = ArrDims(arr)
    If cDims = 0 Then Err.Raise vbErrorInvalidProcedureCall
    If dimension < 1 Or dimension > cDims Then Err.Raise vbErrorSubscriptOutOfRange

    ' The highest array dimension has the smallest memory offset.
    Dim Offset As LongPtr: Offset = vbSizeLong + 2 * vbSizeLong * (cDims - dimension)
    PokeLng ArrPtr(arr) + vbOffsetArrayrgsaBound + Offset, RHS

End Property


Public Property Get MidArr( _
    ByRef arr As Variant, _
    ByVal Start As Long, _
    Optional ByVal Length As Variant _
) As Variant
Attribute MidArr.VB_Description = "Returns a slice from an array / replaces a slice in an array."
' Get MidArr is the array analogue of the VBA.Mid function.

    ' Test the input parameters before manipulating memory.
    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa = vbNullPtr Then
        MidArr = CreateArray(VBA.VarType(arr))
        Exit Property
    ElseIf Start > UBound(arr) Then
        MidArr = CreateArray(VBA.VarType(arr))
        Exit Property
    ElseIf Start < LBound(arr) Then
        Err.Raise vbErrorSubscriptOutOfRange, , "Start is invalid"
    End If

    ' Check the length parameter.
    If VBA.IsMissing(Length) Then
        Length = UBound(arr) - Start + 1
    ElseIf Length > UBound(arr) - Start + 1 Then
        Length = UBound(arr) - Start + 1
    ElseIf Length = 0 Then
        MidArr = CreateArray(VBA.VarType(arr))
        Exit Property
    ElseIf Length < 0 Then
        Err.Raise vbErrorInvalidProcedureCall, , "Length is invalid"
    End If

    ' Copy the array slice by temporarily tweaking the SafeArray.
    With PeekSafeArray(psa)
        If .cDims <> 1 Then Err.Raise vbErrorInvalidProcedureCall, , "Array has multiple dimensions"
        Dim Count As Long: Count = this.MetaData.cElements
        Dim Data As LongPtr: Data = this.MetaData.pvData
        this.MetaData.cElements = 1
        this.MetaData.pvData = psa + vbOffsetArraypvData
        this.Pointers(0) = .pvData + (Start - .lLbound) * .cbElements
        this.MetaData.pvData = psa + vbOffsetArrayrgsaBound
        this.Longs(0) = Length
        MidArr = arr
        this.Longs(0) = .cElements
        this.MetaData.pvData = psa + vbOffsetArraypvData
        this.Pointers(0) = .pvData
        this.MetaData.cElements = Count
        this.MetaData.pvData = Data
    End With

End Property
Public Property Let MidArr( _
    ByRef arr As Variant, _
    ByVal Start As Long, _
    Optional ByVal Length As Variant, _
    ByVal RHS As Variant _
)
' Let MidArr is the array analogue of the Mid statement.
' Keep ByVal for RHS. RHS is destroyed to compensate for reference counting.

    ' Test the input parameters before manipulating memory.
    Dim psa As LongPtr: psa = ArrPtr(arr)
    If psa = vbNullPtr Then
        Exit Property
    ElseIf Start > UBound(arr) Then
        Exit Property
    ElseIf Start < LBound(arr) Then
        Err.Raise vbErrorSubscriptOutOfRange, , "Start is invalid"
    End If
    ' RHS is not checked for being one dimensional.
    Dim source As LongPtr: source = ArrDataPtr(RHS)
    If source = vbNullPtr Then
        Exit Property
    End If

    ' Check the length parameter.
    If VBA.IsMissing(Length) Then
        Length = UBound(RHS) - LBound(RHS) + 1
    ElseIf Length > UBound(RHS) - LBound(RHS) + 1 Then
        Length = UBound(RHS) - LBound(RHS) + 1
    ElseIf Length < 0 Then
        Err.Raise vbErrorInvalidProcedureCall, , "Length is invalid"
    ElseIf Length = 0 Then
        Exit Property
    End If
    ' Limit the length of the copied array slice to the maximum possible.
    If Start + Length - 1 > UBound(arr) Then
        Length = UBound(arr) - Start + 1
    End If

    ' Get the arguments for using the CopyMem function.
    With PeekSafeArray(psa)
        If .cDims <> 1 Then Err.Raise vbErrorInvalidProcedureCall, , "Array has multiple dimensions"
        Dim NBytes As Long: NBytes = Length * .cbElements
        Dim destination As LongPtr: destination = .pvData + (Start - .lLbound) * .cbElements
    End With

    ' Copying raw data ignores object reference counting and may lead to double
    ' BSTR references, both of which are illegal and will give a system crash.
    ' In order to prevent this happening, the destination data is copied to an
    ' garbage array before the destination data is overwritten, and the source
    ' data is zeroed after it is copied. The RHS (source) array is passed ByVal.
    Select Case VBA.VarType(arr)
    Case Is <> VBA.VarType(RHS)
        Err.Raise vbErrorTypeMismatch, , "Array types do not match"
    Case VBA.vbVariant Or VBA.vbArray
        Dim VarGarbage() As Variant: ReDim VarGarbage(0 To Length - 1)
        CopyMem VarPtr(VarGarbage(0)), destination, NBytes
        CopyMem destination, source, NBytes
        ZeroMem source, NBytes
    Case VBA.vbString Or VBA.vbArray
        Dim StrGarbage() As String: ReDim StrGarbage(0 To Length - 1)
        CopyMem VarPtr(StrGarbage(0)), destination, NBytes
        CopyMem destination, source, NBytes
        ZeroMem source, NBytes
    Case VBA.vbObject Or VBA.vbArray
        Dim ObjGarbage() As Object: ReDim ObjGarbage(0 To Length - 1)
        CopyMem VarPtr(ObjGarbage(0)), destination, NBytes
        CopyMem destination, source, NBytes
        ZeroMem source, NBytes
    Case Else
        CopyMem destination, source, NBytes
    End Select

End Property


Public Sub CopyMem( _
    ByVal destination As LongPtr, _
    ByVal source As LongPtr, _
    ByVal NBytes As Long _
)
Attribute CopyMem.VB_Description = "Copies a memory block from a source address to a destination address."
' The CopyMem function uses LSet to copy a memory block from a source address
' to a destination address which is made possible by a hack which tricks VB to
' handle the source and destination memory blocks as strings. A VB String is in
' fact a BSTR which requires six overhead bytes plus 2 bytes for each character.
' A VB String thus always has an even number of bytes. It uses the 4 bytes just
' before the string address to store the number of bytes in the string and it is
' terminated by two null bytes. The actual code requires a combination of LSet
' and Poke/Peek calls to deal with uneven as well as small blocks of memory. No
' provisions are made for the termination null bytes as they seem to be ignored.

    ' Check the validity of the arguments.
    If NBytes <= 0 Then Exit Sub
    If destination = vbNullPtr Then Err.Raise vbErrorInvalidProcedureCall
    If source = vbNullPtr Then Err.Raise vbErrorInvalidProcedureCall

    ' The CopyMemory API is faster above the (approximate) threshold.
    Const Threshold As Long = 10 ^ 6
    If NBytes >= Threshold Then
        CopyMemory ByVal destination, ByVal source, NBytes
        Exit Sub
    End If

    ' Temporary variables are required to copy from source to destination.
    Dim Tmp1 As Byte, Tmp2 As Integer, Tmp4 As Long, Tmp8 As Currency

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = NBytes
    this.MetaData.pvData = source
    ' Copy the last byte if there is an uneven number of bytes.
    If NBytes Mod 2 = 1 Then
        NBytes = NBytes - 1
        Tmp1 = this.Bytes(NBytes)
        this.MetaData.pvData = destination
        this.Bytes(NBytes) = Tmp1
        this.MetaData.pvData = source
    End If
    ' Copy the remaining even number of bytes.
    Select Case NBytes
    Case Is >= 10
        Dim BSTR(0 To 1) As LongPtr
        BSTR(0) = destination + 4
        BSTR(1) = source + 4
        Tmp4 = this.Longs(0)
        this.Longs(0) = NBytes - 4
        this.MetaData.pvData = destination
        this.Longs(0) = NBytes - 4
        this.MetaData.pvData = VarPtr(BSTR(0))
        LSet this.Strings(0) = this.Strings(1)
        this.MetaData.pvData = source
        this.Longs(0) = Tmp4
        this.MetaData.pvData = destination
        this.Longs(0) = Tmp4
    Case 8
        Tmp8 = this.Currencies(0)
        this.MetaData.pvData = destination
        this.Currencies(0) = Tmp8
    Case 6
        Tmp4 = this.Longs(0)
        Tmp2 = this.Integers(2)
        this.MetaData.pvData = destination
        this.Longs(0) = Tmp4
        this.Integers(2) = Tmp2
    Case 4
        Tmp4 = this.Longs(0)
        this.MetaData.pvData = destination
        this.Longs(0) = Tmp4
    Case 2
        Tmp2 = this.Integers(0)
        this.MetaData.pvData = destination
        this.Integers(0) = Tmp2
    End Select
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Sub ZeroMem( _
    ByVal address As LongPtr, _
    ByVal NBytes As Long _
)
Attribute ZeroMem.VB_Description = "Copies null bytes to a memory block at a destination address."
' The ZeroMem function works similar to the CopyMem function. However, it needs
' to generate a Null string of size NBytes\2 which is costly. Therefore, ZeroMem
' switches to the use of the ZeroMemory API at a relatively low threshold.

    ' Check the validity of the arguments.
    If NBytes <= 0 Then Exit Sub
    If address = vbNullPtr Then Err.Raise vbErrorInvalidProcedureCall

    ' The ZeroMemory API is faster above the (approximate) threshold.
    Const Threshold As Long = 10 ^ 4
    If NBytes >= Threshold Then
        ZeroMemory ByVal address, NBytes
        Exit Sub
    End If

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = NBytes
    this.MetaData.pvData = address
    ' Zero the last byte if there is an uneven number of bytes.
    If NBytes Mod 2 = 1 Then
        NBytes = NBytes - 1
        this.Bytes(NBytes) = 0
    End If
    ' Zero the remaining even number of bytes.
    Select Case NBytes
    Case Is >= 10
        Dim BSTR As LongPtr
        BSTR = address + 4
        this.Longs(0) = NBytes - 4
        this.MetaData.pvData = VarPtr(BSTR)
        LSet this.Strings(0) = String$((NBytes - 4) \ 2, VBA.vbNullChar)
        this.MetaData.pvData = address
        this.Longs(0) = 0
    Case 8
        this.Currencies(0) = 0
    Case 6
        this.Longs(0) = 0
        this.Integers(2) = 0
    Case 4
        this.Longs(0) = 0
    Case 2
        this.Integers(0) = 0
    End Select
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Public Function DumpMem( _
    ByVal address As LongPtr, _
    Optional ByVal NBytes As Long = 256, _
    Optional ByVal NBytesPerLine As Long = 16, _
    Optional ByVal NBytesPerDash As Long = 8 _
) As String
Attribute DumpMem.VB_Description = "Returns the hexadecimal dump of a memory block at an address."
' After Karl E. Peterson - see http://www.mvps.org/vb.
' Here the dump is written to a string buffer which avoids intermediate string
' allocations and thus makes the code faster for (larger) memory dumps.

    ' Check the validity of the arguments.
    If NBytes <= 0 Then Exit Function
    If address = vbNullPtr Then Err.Raise vbErrorInvalidProcedureCall

    ' NBytesPerLine and NBytesPerDash are set to multiples of eight.
    If NBytesPerLine <> 16 Then NBytesPerLine = 8 * (NBytesPerLine \ 8)
    If NBytesPerDash <> 8 Then NBytesPerDash = 8 * (NBytesPerDash \ 8)

    ' Load the memory block into a byte array rounded up to multiple of NBytesPerLine.
    Dim NLines As Long: NLines = (NBytes + NBytesPerLine - 1) \ NBytesPerLine
    Dim Bytes() As Byte: ReDim Bytes(0 To NBytesPerLine * NLines - 1) As Byte
    CopyMem VarPtr(Bytes(0)), address, NBytes

    ' Initialize the buffer with 3 extra lines and 1 extra character per line.
    Dim Width As Long: Width = 9 + 6 + 3 * NBytesPerLine + 1 + NBytesPerLine
    Dim Points As String: Points = String$(NBytesPerLine, ".")
    Dim Buffer As String: Buffer = Space$((NLines + 3) * (Width + 1))
    Dim index As Long: index = 1

    ' Write the header to the buffer.
    Mid$(Buffer, index) = String$(Width, "=") & VBA.vbCrLf
    index = index + Width + 1

    ' Write the information line to the buffer.
    Mid$(Buffer, index) = "Memory address = &h" & Hex$(address)
    index = index + Width + 1
    Dim Info As String: Info = "# bytes = " & NBytes & VBA.vbCrLf
    Mid$(Buffer, index - Len(Info) + 1) = Info

    ' Loop and write the memory dump to the buffer.
    Dim n As Long, i As Long
    For n = 0 To NBytes - 1 Step NBytesPerLine
        ' Write the 32-bit memory address.
        Mid$(Buffer, index) = Right$("0000000" & Hex$(address + n), 8)
        index = index + 9

        ' Write the 16-bit memory offset.
        Mid$(Buffer, index) = Right$("000" & Hex$(n), 4)
        index = index + 6

        ' Write the Hex codes.
        For i = n To n + NBytesPerLine - 1
            Mid$(Buffer, index) = Right$("0" & Hex$(Bytes(i)), 2)
            index = index + 3
            If (i + 1) Mod NBytesPerDash = 0 Then
                Mid$(Buffer, index - 1) = "-"
            End If
        Next
        Mid$(Buffer, index - 1) = " "
        index = index + 1

        ' Write the Ascii characters.
        Mid$(Buffer, index) = Points
        For i = n To n + NBytesPerLine - 1
            If Bytes(i) >= 32 And Bytes(i) <= 126 Then
                Mid$(Buffer, index) = Chr$(Bytes(i))
            End If
            index = index + 1
        Next

        ' Write newline.
        Mid$(Buffer, index) = VBA.vbCrLf
        index = index + 1
    Next

    ' Write the footer to the buffer.
    Mid$(Buffer, index) = String$(Width, "=") & VBA.vbCrLf

    ' Erase any byte information in the buffer that is outside the scope.
    n = NBytes Mod NBytesPerLine
    If n <> 0 Then
        ' Number of bytes for which information needs to be erased.
        n = NBytesPerLine - n
        Mid$(Buffer, index - 1 - n) = Space$(n)
        Mid$(Buffer, index - 1 - NBytesPerLine - 1 - 3 * n) = Space$(3 * n)
    End If

    DumpMem = Buffer

End Function


Public Function CloneArray(ByVal address As LongPtr) As Variant
Attribute CloneArray.VB_Description = "Returns an array clone of a SafeArray at a memory address."
' CloneArray can also be used to clone a ParamArray so that the elements of the
' clone remain Byref and can be passed to the next routine which may then change
' the value of one or more of the original ParamArray arguments. The address of
' the SafeArray of a ParamArray is obtained by ArrPtr(ArrVarPtr(args)).

    If address = vbNullPtr Then Err.Raise vbErrorInvalidProcedureCall

    Dim arr As Variant: arr = address
    Dim vt As Integer
    With PeekSafeArray(address)
        ' A ParamArray does not have FADF_HAVEVARTYPE set.
        If (.fFeatures And FADF_HAVEVARTYPE) = FADF_HAVEVARTYPE Then
            vt = PeekLng(address + vbOffsetArrayType)
        End If
        ' A ParamArray does have FADF_VARIANT set.
        ' Try to get the SafeArray type via Advanced Features.
        If vt = 0 Then
            If (.fFeatures And FADF_VARIANT) = FADF_VARIANT Then
                vt = VBA.vbVariant
            ElseIf (.fFeatures And FADF_BSTR) = FADF_BSTR Then
                vt = VBA.vbString
            Else
                Err.Raise vbErrorTypeMismatch, , "Can't determine type of SafeArray."
            End If
        End If
    End With

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = VarPtr(arr)
    this.Integers(0) = vt Or VBA.vbArray
    CloneArray = arr
    this.Integers(0) = VBA.vbEmpty
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data
    PokeInt VarPtr(arr), VBA.vbEmpty

End Function


Public Function CollKeys( _
    ByVal coll As Collection, _
    Optional ByVal base As Long _
) As String()
Attribute CollKeys.VB_Description = "Returns an array containing the keys of all the items in a VB Collection."
' Returns an empty String array for an empty Collection.

    If coll.Count = 0 Then Exit Function
    Dim Keys() As String: ReDim Keys(base To base + coll.Count - 1)

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionHead
    ' Traverse the double linked list and get all the keys.
    Dim i As Long, NodePtr As LongPtr
    For i = base To base + coll.Count - 1
        NodePtr = this.Pointers(0)
        this.MetaData.pvData = NodePtr + vbOffsetCollectionKey
        Keys(i) = this.Strings(0)
        this.MetaData.pvData = NodePtr + vbOffsetCollectionNext
    Next
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data
    CollKeys = Keys

End Function


Public Function CollKey( _
    ByVal coll As Collection, _
    ByVal pos As Long _
) As String
Attribute CollKey.VB_Description = "Returns the key of an item in a VB Collection specified by its position."
' Returns "" for a Collection item that doesn't have a key.

    If coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Collection is empty"
    If pos < 1 Or pos > coll.Count Then Err.Raise vbErrorSubscriptOutOfRange

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    ' Traverse the double linked list depending on the shortest path to the node position.
    this.MetaData.cElements = 1
    If pos <= coll.Count \ 2 Then
        this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionHead
        Dim i As Long, NodePtr As LongPtr
        For i = 1 To pos
            NodePtr = this.Pointers(0)
            this.MetaData.pvData = NodePtr + vbOffsetCollectionNext
        Next
    Else
        this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionTail
        For i = coll.Count To pos Step -1
            NodePtr = this.Pointers(0)
            this.MetaData.pvData = NodePtr + vbOffsetCollectionPrev
        Next
    End If
    this.MetaData.pvData = NodePtr + vbOffsetCollectionKey
    CollKey = this.Strings(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Function CollPos( _
    ByVal coll As Collection, _
    ByVal key As String _
) As Long
Attribute CollPos.VB_Description = "Returns the position of an item in a VB Collection specified by its key."
' Returns 0 if no Collection item with a given key exists.

    If coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Collection is empty"
    If KeyExists(coll, key) = False Then Exit Function

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    ' Traverse the double linked list until a matching key has been found.
    this.MetaData.cElements = 1
    this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionHead
    Dim i As Long, NodePtr As LongPtr
    For i = 1 To coll.Count
        NodePtr = this.Pointers(0)
        this.MetaData.pvData = NodePtr + vbOffsetCollectionKey
        ' Collection Keys are compare as case-insensitive.
        If StrComp(this.Strings(0), key, VBA.vbTextCompare) = 0 Then
            CollPos = i
            Exit For
        End If
        this.MetaData.pvData = NodePtr + vbOffsetCollectionNext
    Next
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Public Property Let CollItem( _
    ByVal coll As Collection, _
    ByVal index As Variant, _
    ByVal RHS As Variant _
)
Attribute CollItem.VB_Description = "Sets the value of an item in a VB Collection specified by its index."
' Works for objects as well as non-objects, like the add method for a Collection.

    If coll.Count = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Collection is empty"

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    ' Traverse the double linked list until a matching item has been found.
    If VBA.VarType(index) = VBA.vbString Then
        If KeyExists(coll, index) = False Then Err.Raise vbErrorInvalidProcedureCall
        this.MetaData.cElements = 1
        this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionHead
        Dim i As Long, NodePtr As LongPtr
        For i = 1 To coll.Count
            NodePtr = this.Pointers(0)
            this.MetaData.pvData = NodePtr + vbOffsetCollectionKey
            If StrComp(this.Strings(0), index, VBA.vbTextCompare) = 0 Then
                Exit For
            End If
            this.MetaData.pvData = NodePtr + vbOffsetCollectionNext
        Next
    Else
        If index < 1 Or index > coll.Count Then Err.Raise vbErrorSubscriptOutOfRange
        this.MetaData.cElements = 1
        If index <= coll.Count \ 2 Then
            this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionHead
            For i = 1 To index
                NodePtr = this.Pointers(0)
                this.MetaData.pvData = NodePtr + vbOffsetCollectionNext
            Next
        Else
            this.MetaData.pvData = ObjPtr(coll) + vbOffsetCollectionTail
            For i = coll.Count To index Step -1
                NodePtr = this.Pointers(0)
                this.MetaData.pvData = NodePtr + vbOffsetCollectionPrev
            Next
        End If
    End If
    this.MetaData.pvData = NodePtr
    If VBA.IsObject(RHS) Then
        Set this.Variants(0) = RHS
    Else
        this.Variants(0) = RHS
    End If
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub InitializeConstruct()
Attribute InitializeConstruct.VB_Description = "Initializes the special construct."
 ' All this arrays fields are pointed to the (shared) this.MetaData field.

    Dim MetaData As LongPtr: MetaData = VarPtr(this.MetaData)
    Dim Delta As Long: Delta = LenB(this) - LenB(this.MetaData)
    Dim Offset As Long: Offset = VBA.Abs(MetaData - VarPtr(this))

    ' The code expects the SafeArray (MetaData) as last field.
    If Offset <> Delta Then Err.Raise vbErrorInternalError

    ' Set the Pointers array variable to the SafeArray (MetaData) field.
    CopyMemory ByVal VarPtrArray(this.Pointers), MetaData, vbSizeLongPtr

    With this.MetaData
        ' Initialize the SafeArray (MetaData) and set the data pointer to this.
        .cDims = 1
        .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        ' The element size is derived from the type of the actual array variable.
        .cbElements = 0
        ' Lock the SafeArray so that it can't be accidentally changed by VB code.
        .cLocks = 1
        .lLbound = 0
        ' The number of elements is set to the number of array variables in TMemory.
        .cElements = Delta \ vbSizeLongPtr
        ' Set the array data pointer to the first array variable in the construct.
        .pvData = VarPtr(this)

        ' Set all the array variables to the (same) SafeArray (MetaData) field.
        Dim i As Long
        For i = 0 To .cElements - 1
            this.Pointers(i) = MetaData
        Next

        ' Protect from automatic garbage collection.
        .cElements = 0
        .pvData = vbNullPtr
    End With

End Sub


Private Function PeekSafeArray(ByVal address As LongPtr) As SAFEARRAY
Attribute PeekSafeArray.VB_Description = "Copies a SafeArray structure located at a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    PeekSafeArray = this.SafeArrays(0)
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Function


Private Sub PokeSafeArray(ByVal address As LongPtr, ByRef value As SAFEARRAY)
Attribute PokeSafeArray.VB_Description = "Copies a SafeArray structure to a memory address."

    If this.MetaData.cDims = 0 Then InitializeConstruct

    Dim Count As Long: Count = this.MetaData.cElements
    Dim Data As LongPtr: Data = this.MetaData.pvData
    this.MetaData.cElements = 1
    this.MetaData.pvData = address
    this.SafeArrays(0) = value
    this.MetaData.cElements = Count
    this.MetaData.pvData = Data

End Sub


Private Function KeyExists(ByVal coll As Collection, ByVal key As String) As Boolean
Attribute KeyExists.VB_Description = "Returns True if a key exists in a VB Collection or False otherwise."
' An error trap is required to check whether a key exists or not.

    On Error GoTo ErrorHandler
    coll.Item key
    KeyExists = True
    Exit Function

ErrorHandler:
    Select Case Err.Number
    Case vbErrorInvalidProcedureCall
        Err.Clear
    Case vbErrorObjectVariableNotSet, vbErrorObjectRequired
        Err.Clear
    Case Else
        Err.Raise Err.Number
    End Select

End Function
