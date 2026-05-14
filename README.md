# vba-dma
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Direct Memory Access standard module for VBA.

Peek/Poke, array utilities, memory block operations, and Collection extensions.

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

## 📄 License

MIT © 2025 Vincent van Geerestein
