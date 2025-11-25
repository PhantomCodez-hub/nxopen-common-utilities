# NX Open Common Utility Library

A reusable, production-style **NX Open Common Utility** built in C# to accelerate Siemens NX automation.

This library groups frequently used operations into a single, structured namespace (`NxCommonUtilities`) and adds **defensive coding** with `try/catch`, logging, and clear error messages.

---

## 1. Overview

### Purpose

- Provide **common helper functions** for daily NX customization work.
- Avoid repeating boilerplate code when:
  - Extruding curves or sketches
  - Performing Boolean subtracts
  - Creating offsets
  - Trimming based on volume
  - Extracting dumb bodies
  - Projecting curves to datum planes
  - Measuring volumes, lengths, distances

### Key Design Choices

- All methods are **static** under `NxCommonUtilities.NxCommonUtility`.
- Grouped into logical regions:
  - Color
  - Extrude
  - Boolean / Body
  - Extract
  - Offset
  - Trim By Volume
  - Measurement / Evaluation
  - Sketch & Selection
  - Datum & Projection
  - General Helpers
- Every public method:
  - Validates input (null / empty checks)
  - Uses `try/catch` and logs error context to the **NX Listing Window**
  - Where reasonable, returns `null` / empty array instead of crashing the session.

---

## 2. How to Integrate in an NX Open Project

### 2.1. Add File

1. Create a folder `Common` or `Utilities` in your NX Open C# project.
2. Add `NxCommonUtility.cs` file.
3. Ensure you reference:
   - `NXOpen.dll`
   - `NXOpen.UF.dll`
   - `NXOpen.Utilities.dll`
   - And others you already use (`NXOpen.BlockStyler`, etc.).

### 2.2. Namespace Usage

```csharp
using NxCommonUtilities;
using NXOpen;
using NXOpen.Features;
