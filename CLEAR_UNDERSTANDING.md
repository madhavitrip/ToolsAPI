# Clear Understanding of Both Endpoints

## EnvelopeBreakage Endpoint [HttpGet("EnvelopeBreakage")]

**What it does:**
- Reads NRData + EnvelopeBreakages table
- Adds extras (Nodal Extra, University Extra, Office Extra)
- Replicates rows based on TotalEnvelope count
- Generates SerialNumber (resets per CatchNo)
- Calculates BookletSerial ranges

**Excel Output (EnvelopeBreaking.xlsx):**
```
SerialNumber | CatchNo | CenterCode | Quantity | EnvQuantity | CenterEnv | TotalEnv | Env | BookletSerial | CourseName | ...
1            | C001    | CC1        | 100      | 50          | 1         | 2        | 1/2 | 1-50          | Math       |
2            | C001    | CC1        | 100      | 50          | 2         | 2        | 2/2 | 51-100        | Math       |
3            | C002    | CC2        | 75       | 75          | 1         | 1        | 1/1 | 101-175       | Science    |
```

**Key Fields:**
- SerialNumber (1, 2, 3... resets per CatchNo)
- EnvQuantity (quantity per envelope)
- CenterEnv (counter per CatchNo-CenterCode)
- TotalEnv (total envelopes for this NRData)
- Env (1/2, 2/2 format)
- BookletSerial (range like 1-50)

---

## Replication Endpoint [HttpGet("Replication")]

**What it does:**
- Reads EnvelopeBreaking.xlsx (from above)
- Removes duplicates
- Applies sorting (CenterSort, NodalSort, RouteSort, ExamDate)
- Implements box-breaking logic (splits by box capacity)
- Calculates Start/End/Serial for each box
- Generates OmrSerial ranges
- Generates InnerBundlingSerial

**Excel Output (BoxBreaking.xlsx):**
```
SerialNumber | CatchNo | CenterCode | Quantity | Start | End | Serial    | TotalPages | BoxNo | OmrSerial | InnerBundlingSerial | CourseName |
1            | C001    | CC1        | 50       | 1     | 5   | 1 to 5    | 250        | 1     | 1-50      | 1                   | Math       |
2            | C001    | CC1        | 50       | 6     | 10  | 6 to 10   | 250        | 2     | 51-100    | 1                   | Math       |
3            | C002    | CC2        | 75       | 11    | 15  | 11 to 15  | 375        | 3     | 101-175   | 2                   | Science    |
```

**Key Fields:**
- SerialNumber (1, 2, 3... continuous)
- Start/End (envelope serial range)
- Serial (formatted "1 to 5")
- TotalPages (Quantity * Pages per NRData)
- BoxNo (box assignment)
- OmrSerial (OMR serial range)
- InnerBundlingSerial (bundling group)

---

## The Key Difference

**Same NRData (C001) produces DIFFERENT data:**

| Field | EnvelopeBreakage | Replication |
|-------|-----------------|-------------|
| SerialNumber | 1, 2 (resets per CatchNo) | 1, 2 (continuous) |
| Quantity | 100 (original) | 50, 50 (split by box) |
| EnvQuantity | 50 (per envelope) | - (not in Replication) |
| CenterEnv | 1, 2 (counter) | - (not in Replication) |
| Start/End | - (not in EnvelopeBreakage) | 1-5, 6-10 (box ranges) |
| Serial | - (not in EnvelopeBreakage) | "1 to 5", "6 to 10" |
| TotalPages | - (not in EnvelopeBreakage) | 250, 250 |
| BoxNo | - (not in EnvelopeBreakage) | 1, 2 |
| OmrSerial | - (not in EnvelopeBreakage) | 1-50, 51-100 |

---

## Solution: TWO Separate Tables

### Table 1: EnvelopeBreakingResult
Store data from **EnvelopeBreakage endpoint**:
- ProjectId, NrDataId
- CatchNo, CenterCode, Quantity
- EnvQuantity, CenterEnv, TotalEnv, Env
- SerialNumber, BookletSerial
- CourseName, ExamDate, ExamTime, NodalCode, Route, etc.

### Table 2: BoxBreakingResult
Store data from **Replication endpoint**:
- ProjectId, NrDataId
- CatchNo, CenterCode, Quantity
- Start, End, Serial, TotalPages
- BoxNo, OmrSerial, InnerBundlingSerial
- SerialNumber, CourseName, ExamDate, ExamTime, NodalCode, Route, etc.

---

## Why This Works

✅ Each endpoint stores its own output independently
✅ No conflicts or overwrites
✅ Same NRData can have different values in both tables (correct!)
✅ No duplication issues
✅ Both Excel files match the database exactly
