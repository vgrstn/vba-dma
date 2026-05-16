# vba-dma
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA standard module for Direct Memory Access — Peek/Poke, and Collection extensions, array utilities, memory block operations.

Uses a **SafeArray construct** for fast memory access without extra stack frames, inspired by [Christian Buse's VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools). Unlike LibMemory, a single shared construct handles all methods with nested call support via save/restore of the data pointer.

---

## 📦 Features

- **Peek/Poke** — read and write all intrinsic VBA types at any memory address
- **Array utilities** — pointer, dimensions, element size, static/empty/allocated checks, base, element access, and array slicing
- **Memory block ops** — `CopyMem`, `ZeroMem`, `DumpMem` with automatic API fallback above thresholds
- **Collection extensions** — retrieve keys, positions, and update items by index or key
- x86 / x64 compatible via `LongPtr` and `#If Win64` throughout
- Pure VBA, Rubberduck-friendly annotations

---

## 📁 Files

| File | Description |
|---|---|
| `DMA.bas` | Source file with [Rubberduck](https://rubberduckvba.com/) annotations |
| `DMA_WithAttributes.bas` | Ready-to-import version with VB attributes baked in — no Rubberduck required |

---

## ⚙️ Public Interface

### Peek / Poke

| Member | Returns | Description |
|---|---|---|
| `PeekBool(address)` | `Boolean` | Reads a Boolean at a memory address |
| `PeekByte(address)` | `Byte` | Reads a Byte at a memory address |
| `PeekCur(address)` | `Currency` | Reads a Currency at a memory address |
| `PeekDbl(address)` | `Double` | Reads a Double at a memory address |
| `PeekInt(address)` | `Integer` | Reads an Integer at a memory address |
| `PeekLng(address)` | `Long` | Reads a Long at a memory address |
| `PeekLngLng(address)` *(x64)* | `LongLong` | Reads a LongLong at a memory address |
| `PeekPtr(address)` | `LongPtr` | Reads a LongPtr at a memory address |
| `PeekSng(address)` | `Single` | Reads a Single at a memory address |
| `PeekVar(address)` | `Variant` | Reads a Variant at a memory address |
| `PeekStr(address)` | `String` | Reads a String (BSTR) at a memory address |
| `PeekObj(address)` | `Object` | Reads an Object reference at a memory address |
| `PokeBool(address, value)` | — | Writes a Boolean to a memory address |
| `PokeByte(address, value)` | — | Writes a Byte to a memory address |
| `PokeCur(address, value)` | — | Writes a Currency to a memory address |
| `PokeDbl(address, value)` | — | Writes a Double to a memory address |
| `PokeInt(address, value)` | — | Writes an Integer to a memory address |
| `PokeLng(address, value)` | — | Writes a Long to a memory address |
| `PokeLngLng(address, value)` *(x64)* | — | Writes a LongLong to a memory address |
| `PokePtr(address, value)` | — | Writes a LongPtr to a memory address |
| `PokeSng(address, value)` | — | Writes a Single to a memory address |
| `PokeVar(address, value)` | — | Writes a Variant to a memory address |
| `PokeStr(address, value)` | — | Writes a String (BSTR) to a memory address |
| `PokeObj(address, value)` | — | Writes an Object reference to a memory address |

### Array Utilities

| Member | Returns | Description |
|---|---|---|
| `ArrPtr(arr)` | `LongPtr` | Address of the underlying SafeArray |
| `ArrVarPtr(arr)` | `LongPtr` | Address of the array variable |
| `ArrDataPtr(arr)` | `LongPtr` | Address of the stored data |
| `ArrDims(arr)` | `Integer` | Number of dimensions (0 if unallocated) |
| `ArrElemSize(arr)` | `Long` | Element size in bytes |
| `IsArrayAllocated(arr)` | `Boolean` | True if the array has an allocated SafeArray |
| `IsArrayEmpty(arr)` | `Boolean` | True if unallocated or zero elements |
| `IsArrayStatic(arr)` | `Boolean` | True if the array is static |
| `ArrBase(arr [, dim])` *(Let)* | — | Sets the lower bound of a dimension |
| `ArrElem(arr, indices)` *(Get/Let/Set)* | `Variant` | Gets or sets an element by index vector |
| `MidArr(arr, start [, length])` *(Get/Let)* | `Variant` | Gets or sets an array slice |
| `CreateArray(vt [, dims])` | `Variant` | Creates a typed array at runtime |
| `CloneArray(address)` | `Variant` | Clones a SafeArray from a memory address |

### Memory Block Operations

| Member | Description |
|---|---|
| `CopyMem(dest, src, NBytes)` | Copies a memory block; uses `RtlMoveMemory` above 1 MB |
| `ZeroMem(address, NBytes)` | Zeroes a memory block; uses `RtlZeroMemory` above 10 KB |
| `DumpMem(address [, NBytes])` | Returns a formatted hexadecimal memory dump |
| `OffsetPtr(address, offset)` | Safe signed pointer arithmetic |

### Collection Extensions

| Member | Returns | Description |
|---|---|---|
| `CollKeys(coll [, base])` | `String()` | All keys in the Collection as an array |
| `CollKey(coll, pos)` | `String` | Key of the item at position `pos` |
| `CollPos(coll, key)` | `Long` | Position of the item with `key` (0 if not found) |
| `CollItem(coll, index)` *(Let)* | — | Assigns a new value to an item by index or key |

### Public Constants

| Constant | Value | Description |
|---|---|---|
| `vbNullPtr` | `0` | Null pointer (missing from VBA) |
| `vbLongPtr` | `vbLong` / `vbLongLong` | VarType constant for `LongPtr` (missing from VBA) |

---

## 🚀 Quick Start

```vb
' Read/write a Long at a known address
Dim x As Long: x = 42
PokeLng VarPtr(x), 99
Debug.Print PeekLng(VarPtr(x))     ' -> 99

' Array slice (MidArr)
Dim a(1 To 5) As Long
a(1) = 10: a(2) = 20: a(3) = 30
Dim s As Variant
s = MidArr(a, 2, 2)               ' -> (20, 30)

' Collection key access
Dim c As New Collection
c.Add "hello", "greeting"
c.Add "world", "subject"
Debug.Print CollKey(c, 1)         ' -> "greeting"
Debug.Print CollPos(c, "subject") ' -> 2
CollItem(c, "greeting") = "hi"
Debug.Print c("greeting")         ' -> "hi"
```

---

## ⏱️ Performance

### Array procedures (ms)

| Method | `Long(100)` | `Long()` | `Variant(100)` | `Variant()` |
|---|---|---|---|---|
| `ArrPtr` | 0.00009 | 0.00009 | 0.00009 | 0.00009 |
| `ArrVarPtr` | 0.00008 | 0.00008 | 0.00008 | 0.00008 |
| `ArrDataPtr` | 0.00018 | 0.00014 | 0.00013 | 0.00015 |
| `ArrDims` | 0.00022 | 0.00013 | 0.00014 | 0.00013 |

### `MidArr` vs `For` loop — `Long(1 To N)` (ms)

| Length | Get MidArr | For loop | Let MidArr | For loop |
|---|---|---|---|---|
| 0 | 0.00056 | 0.00363 | 0.00226 | 0.00445 |
| 10 | 0.00057 | 0.00372 | 0.00267 | 0.00475 |
| 100 | 0.00057 | 0.00624 | 0.00263 | 0.00858 |
| 1,000 | 0.00061 | 0.03157 | 0.00282 | 0.04593 |
| 10,000 | 0.00160 | 0.30272 | 0.00494 | 0.43832 |
| 100,000 | 0.01241 | 2.97955 | 0.04005 | 4.33674 |

### `CopyMem` vs `CopyMemory` API (ms)

| NBytes | CopyMem | CopyMemory |
|---|---|---|
| 10 | 0.00035 | 0.00152 |
| 100 | 0.00032 | 0.00155 |
| 1,000 | 0.00037 | 0.00161 |
| 10,000 | 0.00053 | 0.00169 |
| 100,000 | 0.00275 | 0.00302 |
| 1,000,000 | 0.06700 | 0.06945 |
| 10,000,000 | 0.70597 | 0.68546 ← break even |

### `ZeroMem` vs `ZeroMemory` API (ms)

| NBytes | ZeroMem | ZeroMemory |
|---|---|---|
| 10 | 0.00021 | 0.00185 |
| 100 | 0.00021 | 0.00163 |
| 1,000 | 0.00034 | 0.00162 |
| 10,000 | 0.00153 | 0.00172 ← break even |
| 100,000 | 0.01506 | 0.00152 |
| 1,000,000 | 0.19999 | 0.01422 |
| 10,000,000 | 73.39400 | 0.73259 |

### Collection extensions (ms) — item at position `Count\2`

| Count | `.Item i` | `.Item key` | `CollKey` | `CollPos` | `CollKeys` |
|---|---|---|---|---|---|
| 10 | 0.0001 | 0.0002 | 0.0004 | 0.0011 | 0.0011 |
| 100 | 0.0001 | 0.0002 | 0.0006 | 0.0049 | 0.0094 |
| 1,000 | 0.0011 | 0.0005 | 0.0042 | 0.0449 | 0.0847 |
| 10,000 | 0.0128 | 0.0008 | 0.0360 | 0.4129 | 0.8163 |
| 100,000 | 0.1773 | 0.0010 | 0.3288 | 4.2147 | 9.8930 |

---

## 🧠 How it works

All Peek and Poke operations share a single module-level `CONSTRUCT` UDT containing twelve typed array fields and one `SAFEARRAY` struct (`MetaData`) at the end. `InitializeConstruct` makes every array field point to that single `MetaData` SafeArray. To read or write a value at an arbitrary address, two fields are set and the relevant typed array is accessed:

```vb
this.MetaData.cElements = 1
this.MetaData.pvData = address
value = this.Longs(0)            ' reads 4 bytes
```

No extra stack frame, no helper function, no kernel call. VBA's native array indexing uses the declared element type to determine stride — `this.Bytes(0)` reads 1 byte, `this.Longs(0)` reads 4, `this.Variants(0)` reads 16 or 24 — regardless of the `cbElements` field in the SafeArray. This is the key insight that makes the SafeArray construct faster than both the `CopyMemory` API (kernel transition overhead) and the Variant ByRef method (extra stack frame per access).

---

## 🧠 Implementation notes

### `CONSTRUCT` and `MetaData` — shared SafeArray

```vb
Private Type CONSTRUCT
    Bytes()      As Byte
    Booleans()   As Boolean
    Currencies() As Currency
    Doubles()    As Double
    Integers()   As Integer
    Longs()      As Long
    Objects()    As Object
    Pointers()   As LongPtr
    Singles()    As Single
    Strings()    As String
    Variants()   As Variant
    SafeArrays() As SAFEARRAY   ' copies a full SAFEARRAY struct at once
    MetaData     As SAFEARRAY   ' MUST be last — the single shared SafeArray
End Type
```

`MetaData` must be the last field. `InitializeConstruct` verifies this at runtime with `If Offset <> Delta Then Err.Raise vbErrorInternalError`. All eleven array fields (`Bytes` through `SafeArrays`) are non-allocated; `InitializeConstruct` makes each one point to `MetaData` so they all share the same underlying storage.

### `InitializeConstruct` — bootstrap

```vb
CopyMemory ByVal VarPtrArray(this.Pointers), MetaData, vbSizeLongPtr
```

`VarPtrArray` (an alias for `VarPtr` in VBE7) returns the address of the `Pointers()` array variable — not `VarPtr(this.Pointers)`, which would be the UDT field address. Writing `MetaData` (the address of the `MetaData` struct) into that slot makes `this.Pointers` point to the shared `MetaData` SafeArray. With `cElements = Delta \ vbSizeLongPtr` and `pvData = VarPtr(this)`, the loop then uses `this.Pointers(i)` to write `MetaData` into every other array variable slot in the UDT — all fields now share the same SafeArray.

```vb
.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
.cLocks = 1
.cbElements = 0
```

`FADF_AUTO` marks the array as stack-allocated (not heap), `FADF_FIXEDSIZE` prevents resizing, and `cLocks = 1` prevents OLE automation from destroying it. `cbElements = 0` is correct because VBA's native array indexing ignores this field for typed arrays — it uses the statically-compiled stride for the declared element type.

The idle state at exit (`cElements = 0`, `pvData = vbNullPtr`) prevents VBA's garbage collector from scanning or freeing the array data when the module unloads.

### Static element size — why `cbElements` doesn't matter

When VBA compiles `this.Longs(0)`, it emits code that knows the element size is 4 bytes from the declared type `As Long`. The access is `pvData + 0 * 4`. The `cbElements` field in the SafeArray is used only by OLE automation APIs (`SafeArrayGetElement`, etc.) — not by VBA's native array subscript operator. This is what makes the shared construct work: redirect `pvData` to any address, then read the correct width by choosing the right typed array field.

### Save/restore — nested call support

Every method that uses the construct saves and restores the two live fields:

```vb
Dim Count As Long: Count = this.MetaData.cElements
Dim Data As LongPtr: Data = this.MetaData.pvData
this.MetaData.cElements = 1
this.MetaData.pvData = address
' ... access ...
this.MetaData.cElements = Count
this.MetaData.pvData = Data
```

This allows methods that call other Peek/Poke methods — for example `ArrDataPtr` → `ArrPtr` → checks `this.Integers(0)` → then `ArrDataPtr` calls `PeekPtr` — to all share the one construct without corruption. Each frame saves and restores to whatever the caller had, not to a fixed idle state.

### `PeekStr` / `PokeStr` — double indirection for BSTR

String variables are BSTR pointers — the variable holds a pointer to the string data, not the data itself. `PeekStr` and `PokeStr` take a `StrPtr`-style address (the BSTR pointer value) and use `VarPtr(address)` to point the construct at the stack copy of that pointer:

```vb
this.MetaData.pvData = VarPtr(address)   ' stack copy holds the BSTR pointer value
PeekStr = this.Strings(0)               ' VBA follows the pointer to the string data
```

`PokeStr` uses `LSet` to copy bytes in-place within the existing BSTR allocation — it does not change the string length or reallocate. This is correct for overwriting a live string in memory without leaking the original.

### `PeekObj` / `PokeObj` — object reference semantics

`PeekObj` works the same way as `PeekStr`: `pvData = VarPtr(address)` makes `this.Objects(0)` follow the stack copy of the object pointer, which is then read via `Set PeekObj = this.Objects(0)` — triggering a proper COM `AddRef`.

`PokeObj` is different: `pvData = address` points directly at the target slot. `Set this.Objects(0) = value` writes through the SafeArray into unmanaged memory while still calling `AddRef`/`Release` correctly — appropriate when overwriting an object pointer field in a struct.

### `ArrVarPtr` vs `ArrPtr` — two levels of indirection

A VBA array variable is a pointer to a SafeArray (`**SAFEARRAY` semantics). A `Variant` wrapping an array may or may not carry `VT_BYREF`:

```vb
this.MetaData.pvData = VarPtr(arr)
If (this.Integers(0) And VT_BYREF) = VT_BYREF Then
    ' Typed array variable: Variant data field holds a pointer to the array variable,
    ' which itself holds the SafeArray pointer — two dereferences needed.
    this.MetaData.pvData = this.MetaData.pvData + vbOffsetVariantData
    ArrVarPtr = this.Pointers(0)     ' address of the array variable
Else
    ' Variant/Array: Variant data field IS the array variable slot.
    ArrVarPtr = this.MetaData.pvData + vbOffsetVariantData
End If
```

`ArrPtr` inlines this same logic and adds one more dereference to reach the SafeArray itself, intentionally avoiding a call to `ArrVarPtr` to eliminate the extra stack frame. `ArrDataPtr` calls `ArrPtr` then `PeekPtr` — two nested save/restore cycles, both safe due to the save/restore pattern.

### `CreateArray` — `SafeArrayCreate` with five calling conventions

`SafeArrayCreate(vt, cDims, VarPtr(rgsabound(0)))` allocates the SafeArray on the COM heap and returns a `LongPtr`. The return value of `CreateArray` is a `Variant`; VBA initially stores the raw pointer as an integer. `PokeInt` then overwrites the Variant type field to `vt Or vbArray`, making VBA treat it as an allocated typed array:

```vb
CreateArray = SafeArrayCreate(vt, cDims, VarPtr(rgsabound(0)))
PokeInt VarPtr(CreateArray), vt Or VBA.vbArray
```

The empty-array fast path writes only the type field without allocating a SafeArray — a valid unallocated typed array that can be allocated later with `ReDim`.

Five `rgsabound` layouts are accepted via `Select Case UBound(dims)`:

| Case | Arguments | Meaning |
|---|---|---|
| Missing `dims` | — | Unallocated typed array |
| `UBound = 0`, array arg | `(rgsabound)` | Raw `cElements/lLbound` vector |
| `UBound = 1` | `(cDims, cElements)` | Uniform bound across all dims |
| `UBound = cDims` | `(cDims, e1, …, en)` | Per-dim element counts, zero lower bounds |
| `UBound = 2*cDims` | `(cDims, e1, lb1, …, en, lbn)` | Full spec with lower bounds |

### `ArrBase` — reverse dimension ordering in SafeArrays

```vb
Dim Offset As LongPtr: Offset = vbSizeLong + 2 * vbSizeLong * (cDims - dimension)
PokeLng ArrPtr(arr) + vbOffsetArrayrgsaBound + Offset, RHS
```

SafeArray `rgsaBound` stores dimension bounds in reverse order — the innermost (highest-indexed) dimension is at the lowest memory offset. Dimension `cDims` is at `rgsaBound + 0`; dimension 1 is at `rgsaBound + 2 * vbSizeLong * (cDims - 1)`. The formula skips `vbSizeLong` bytes past the `cElements` Long within each bound pair to reach the `lLbound` field.

### `MidArr` Get — O(1) SafeArray surgery

```vb
With PeekSafeArray(psa)
    ' Redirect pvData forward by Start elements:
    this.Pointers(0) = .pvData + (Start - .lLbound) * .cbElements
    ' Shrink cElements to Length:
    this.Longs(0) = Length
    MidArr = arr          ' VBA copies the now-sliced SafeArray
    ' Restore:
    this.Longs(0) = .cElements
    this.Pointers(0) = .pvData
End With
```

`PeekSafeArray` snapshots the original SafeArray fields into a local `SAFEARRAY` UDT before mutation, so `.pvData`, `.lLbound`, and `.cbElements` remain available for the restore step even after the live SafeArray has been altered. The copy itself is O(length), but the setup is O(1) — no intermediate allocation.

### `MidArr` Let — garbage array for managed types

Raw `CopyMem` of Variant, String, or Object arrays would duplicate live COM/BSTR references without calling `AddRef`/`SysAllocString`, causing double-frees when either copy is collected. The garbage-array pattern absorbs the destination's existing references before overwriting:

```vb
Dim VarGarbage() As Variant: ReDim VarGarbage(0 To Length - 1)
CopyMem VarPtr(VarGarbage(0)), destination, NBytes  ' garbage now owns old dest refs
CopyMem destination, source, NBytes                  ' destination gets new data
ZeroMem source, NBytes                               ' zero RHS so scope exit won't double-free
```

`VarGarbage` goes out of scope and VBA's GC releases each element correctly. `source` is zeroed so that when `RHS` (passed `ByVal`) goes out of scope, VBA finds null references and skips the release. For unmanaged scalars (Long, Integer, etc.), a direct `CopyMem` suffices.

### `CopyMem` — LSet/BSTR trick with adaptive threshold

A VB String is a BSTR: four bytes of length prefix, followed by character data. `CopyMem` exploits this layout by pointing two fake BSTR pointers at the destination and source:

```vb
BSTR(0) = destination + 4   ' fake BSTR at destination, skipping length prefix
BSTR(1) = source + 4        ' fake BSTR at source
this.Longs(0) = NBytes - 4  ' write length prefix at both ends
this.MetaData.pvData = VarPtr(BSTR(0))
LSet this.Strings(0) = this.Strings(1)   ' bulk copy via LSet
```

`LSet` on two same-size strings copies bytes in place without reallocation. The four-byte length prefix and any trailing odd byte are handled with direct typed reads and writes. Below 10 bytes, LSet is bypassed in favour of direct scalar copies (8-, 6-, 4-, 2-byte cases). Above 10⁶ bytes, `RtlMoveMemory` is faster (hardware-optimised memcpy) so the API is used instead.

### `ZeroMem` — `String$` trick with a lower threshold

```vb
this.Longs(0) = NBytes - 4
this.MetaData.pvData = VarPtr(BSTR)
LSet this.Strings(0) = String$((NBytes - 4) \ 2, VBA.vbNullChar)
this.MetaData.pvData = address
this.Longs(0) = 0
```

`String$(n, vbNullChar)` allocates a fresh VBA string of n null characters; `LSet` writes it to the fake BSTR target. Because VBA must allocate this string on every call, the overhead grows with block size — the crossover with `RtlZeroMemory` arrives much earlier than for `CopyMem`, at 10⁴ bytes rather than 10⁶. The timing tables confirm both thresholds empirically.

### `DumpMem` — pre-allocated string buffer

```vb
Dim Buffer As String: Buffer = Space$((NLines + 3) * (Width + 1))
```

The entire output string is allocated in one shot, then filled with `Mid$` assignments. VBA's string concatenation (`&`) allocates a new BSTR on every operation — for a 256-byte dump that would be hundreds of allocations. Writing into a fixed-size buffer via `Mid$` is O(1) per write with no intermediate allocations. After the loop, any byte positions beyond `NBytes` in the final line are overwritten with spaces to erase the pre-allocated padding.

### Collection extensions — doubly-linked list traversal

The VB Collection stores items in a doubly-linked list whose node layout (reconstructed by memory inspection) is:

```
Node: [Item: Variant (16/24 bytes)][Key: LongPtr→BSTR][Prev: LongPtr][Next: LongPtr]
```

Setting `pvData = NodePtr + vbOffsetCollectionKey` and reading `this.Strings(0)` follows the node's BSTR pointer directly — VBA treats the `LongPtr` at that offset as a BSTR pointer and reads the string. A node without a key has `vbNullPtr` at the key offset, so `this.Strings(0)` returns `""`.

`CollKey` traverses from the head if `pos ≤ Count \ 2`, otherwise from the tail, halving worst-case traversal depth. `CollPos` has no equivalent optimisation because each node must be examined regardless of position. `CollKeys` allocates the result `String()` up front and fills it in one forward pass.

### `OffsetPtr` — XOR-based signed arithmetic

```vb
Const SignBitMask As LongPtr = -(2 ^ (8 * vbSizeLongPtr - 1))
OffsetPtr = (address Xor SignBitMask) + Offset Xor SignBitMask
```

Flipping the sign bit before and after the addition converts the signed `LongPtr` to behave as unsigned for the arithmetic, preventing overflow when addresses exceed `&H7FFFFFFF` (x86) or `&H7FFFFFFFFFFFFFFF` (x64). For ordinary pointer offsets within a single object, plain `+` would suffice — `OffsetPtr` is provided for the edge cases where addresses in the high half of the address space are involved.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
