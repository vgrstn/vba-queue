# vba-queue
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Generic Queue based on an encapsulated VB Collection

A lightweight, **Collection-backed FIFO queue** for VBA with:
- O(1) enqueue at the **rear** and dequeue from the **front** (position `1`)
- Safe **enumeration** (`For Each`) via a proper COM enumerator
- Clear error semantics (`Peek` / `Dequeue` on empty queue raise error 5)
- Zero dependencies (pure VBA)

---

## 📁 Files

| File | Description |
|---|---|
| `Queue.cls` | Source file with [Rubberduck](https://rubberduckvba.com/) annotations (`'@Description`, `'@DefaultMember`, `'@Enumerator`) |
| `Queue_WithAttributes.cls` | Ready-to-import version with VB attributes baked in — no Rubberduck required |

Both files are identical in behaviour. Import `Queue_WithAttributes.cls` if you are not using Rubberduck.

---

## 📦 Features

- **Fast enqueue/dequeue** (front is position `1`)
- **Enumeration**: `For Each item In Queue` (via hidden `[_NewEnum]`)
- **Utility export**: `Items([base])` returns a 0- or 1-based array copy
- Pure VBA, no external references, Rubberduck-friendly annotations

---

## ⚙️ Public Interface

| Member               | Type       | Description |
|---------------------|------------|-------------|
| `Enqueue(Item)`      | `Sub`      | Adds an item to the **rear** of the queue. |
| `Dequeue()`          | `Function` | Returns **and removes** the front item. Raises error 5 if empty. |
| `Peek` *(Default)*   | `Property` | Returns the front item **without** removing it. Raises error 5 if empty. |
| `Count`              | `Property` | Number of items. |
| `IsEmpty`            | `Property` | `True` if empty, else `False`. |
| `Clear`              | `Sub`      | Removes all items. |
| `Items([base])`      | `Function` | Returns all items as a `Variant()` array; `base` default = `0`. |
| `For Each`           | Enumerator | Iterates **front → rear** (don't mutate during enumeration). |

**Error behavior**  
- Empty queue on `Peek` / `Dequeue` raises **`vbErrorInvalidProcedureCall (=5)`** with source `"Queue.Peek"` or `"Queue.Dequeue"`.

---

## 🚀 Quick Start

```vb
Dim q As New Queue

q.Enqueue "alpha"
q.Enqueue "beta"
Debug.Print q.Peek        ' -> alpha  (front)
Debug.Print q.Dequeue     ' -> alpha  (removed)
Debug.Print q.Dequeue     ' -> beta   (removed)
Debug.Print q.IsEmpty     ' -> True
```

---

## ⏱️ Performance

Timings (ms) for one `Enqueue` + one `Dequeue`, measured on Windows x64:

| # | Count  | vba-queue | System.Collections.Queue |
|---|--------|-----------|--------------------------|
| 1 | 10      | 0.00045   | 0.00518                  |
| 2 | 100     | 0.00045   | 0.00516                  |
| 3 | 1,000   | 0.00045   | 0.00523                  |
| 4 | 10,000  | 0.00045   | 0.00340                  |
| 5 | 100,000 | 0.00045   | 0.00338                  |

Performance is consistent regardless of queue size. `System.Collections.Queue` uses late binding, which explains its relatively poor performance.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
