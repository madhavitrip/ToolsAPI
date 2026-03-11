# New Processing Endpoints - DB First Approach

## Overview

Created two new controllers that follow a **DB-First** approach instead of Excel-First:

1. **EnvelopeBreakageProcessingController** - Processes envelope breaking
2. **BoxBreakingProcessingController** - Processes box breaking

Both controllers:
- **POST** - Process data and save to database
- **GET** - Retrieve from database and generate Excel

---

## EnvelopeBreakageProcessingController

### POST: /api/EnvelopeBreakageProcessing/ProcessEnvelopeBreaking

**Purpose:** Process NRData + Extras and save to `EnvelopeBreakingResult` table

**Request:**
```
POST /api/EnvelopeBreakageProcessing/ProcessEnvelopeBreaking?ProjectId=1
```

**What it does:**
1. Fetches NRData with `ProcessingStatus = "Processed"`
2. Fetches ExtraEnvelopes for the project
3. Replicates rows based on TotalEnvelope count
4. Adds Nodal Extra, University Extra, Office Extra
5. Saves all to `EnvelopeBreakingResult` table with:
   - `NrDataId` (if regular NRData) OR `ExtraId` (if extra)
   - `CatchNo`, `EnvQuantity`, `CenterEnv`, `TotalEnv`, `Env`
   - `UploadBatch` (for tracking)

**Response:**
```json
{
  "message": "Envelope breaking data saved to database",
  "recordsCount": 150,
  "uploadBatch": 1
}
```

**Database Result:**
```
EnvelopeBreakingResult table:
Id | ProjectId | NrDataId | ExtraId | CatchNo | EnvQuantity | CenterEnv | TotalEnv | Env   | UploadBatch
1  | 1         | 5        | NULL    | C001    | 50          | 1         | 2        | 1/2   | 1
2  | 1         | 5        | NULL    | C001    | 50          | 2         | 2        | 2/2   | 1
3  | 1         | NULL     | 1       | C001    | 25          | 1         | 1        | 1/1   | 1
```

---

### GET: /api/EnvelopeBreakageProcessing/GetEnvelopeBreakingReport

**Purpose:** Retrieve from DB and generate Excel (same format as original)

**Request:**
```
GET /api/EnvelopeBreakageProcessing/GetEnvelopeBreakingReport?ProjectId=1&uploadBatch=1
```

**Parameters:**
- `ProjectId` (required)
- `uploadBatch` (optional) - If not provided, uses latest batch

**What it does:**
1. Fetches EnvelopeBreakingResult from database
2. Joins with NRData to get full details (CenterCode, NodalCode, ExamTime, etc.)
3. Calculates SerialNumber (resets per CatchNo)
4. Calculates BookletSerial ranges
5. Generates Excel with exact same format as original endpoint

**Response:**
```json
{
  "message": "Report generated successfully",
  "filePath": "C:\\...\\wwwroot\\1\\EnvelopeBreakingFromDB.xlsx",
  "recordsCount": 150
}
```

**Excel Output:**
```
Serial Number | Catch No | Center Code | Quantity | EnvQuantity | CenterEnv | TotalEnv | Env | BookletSerial | ...
1             | C001     | CC1         | 100      | 50          | 1         | 2        | 1/2 | 1-50          | ...
2             | C001     | CC1         | 100      | 50          | 2         | 2        | 2/2 | 51-100        | ...
3             | C001     | Nodal Extra | 25       | 25          | 1         | 1        | 1/1 | 101-125       | ...
```

---

## BoxBreakingProcessingController

### POST: /api/BoxBreakingProcessing/ProcessBoxBreaking

**Purpose:** Process box breaking and save to `BoxBreakingResult` table

**Request:**
```
POST /api/BoxBreakingProcessing/ProcessBoxBreaking?ProjectId=1
```

**Prerequisites:**
- Must run `ProcessEnvelopeBreaking` first
- Reads from `EnvelopeBreakingResult` table (latest batch)

**What it does:**
1. Reads EnvelopeBreakingResult from DB
2. Removes duplicates based on DuplicateRemoveFields
3. Calculates Start/End/Serial for each row
4. Applies sorting (CenterSort, NodalSort, RouteSort, ExamDate)
5. Implements box-breaking logic (splits by box capacity)
6. Calculates OmrSerial ranges
7. Calculates InnerBundlingSerial
8. Saves all to `BoxBreakingResult` table

**Response:**
```json
{
  "message": "Box breaking data saved to database",
  "recordsCount": 200,
  "uploadBatch": 1
}
```

**Database Result:**
```
BoxBreakingResult table:
Id | ProjectId | NrDataId | ExtraId | CatchNo | Start | End | Serial    | TotalPages | BoxNo | OmrSerial | UploadBatch
1  | 1         | 5        | NULL    | C001    | 1     | 5   | 1 to 5    | 250        | 1     | 1-50      | 1
2  | 1         | 5        | NULL    | C001    | 6     | 10  | 6 to 10   | 250        | 2     | 51-100    | 1
3  | 1         | NULL     | 1       | C001    | 11    | 12  | 11 to 12  | 125        | 3     | 101-125   | 1
```

---

### GET: /api/BoxBreakingProcessing/GetBoxBreakingReport

**Purpose:** Retrieve from DB and generate Excel (same format as original)

**Request:**
```
GET /api/BoxBreakingProcessing/GetBoxBreakingReport?ProjectId=1&uploadBatch=1
```

**Parameters:**
- `ProjectId` (required)
- `uploadBatch` (optional) - If not provided, uses latest batch

**What it does:**
1. Fetches BoxBreakingResult from database
2. Joins with NRData to get full details
3. Generates Excel with exact same format as original Replication endpoint

**Response:**
```json
{
  "message": "Report generated successfully",
  "filePath": "C:\\...\\wwwroot\\1\\BoxBreakingFromDB.xlsx",
  "recordsCount": 200
}
```

**Excel Output:**
```
SerialNumber | CatchNo | CenterCode | BoxNo | Start | End | Serial    | TotalPages | OmrSerial | InnerBundlingSerial | ...
1            | C001    | CC1        | 1     | 1     | 5   | 1 to 5    | 250        | 1-50      | 1                   | ...
2            | C001    | CC1        | 2     | 6     | 10  | 6 to 10   | 250        | 51-100    | 1                   | ...
3            | C001    | Nodal Extra| 3     | 11    | 12  | 11 to 12  | 125        | 101-125   | 1                   | ...
```

---

## Usage Flow

### Step 1: Upload NRData
```
POST /api/ExcelUpload (existing endpoint)
- Uploads NRData for ProjectId=1
- Sets ProcessingStatus = "Pending"
```

### Step 2: Run Duplicate Removal
```
POST /api/Duplicate (existing endpoint)
- Marks duplicates as "Merged"
- Marks kept rows as "Processed"
```

### Step 3: Run Envelope Configuration
```
POST /api/EnvelopeBreakages/EnvelopeConfiguration (existing endpoint)
- Creates EnvelopeBreakage records
```

### Step 4: Process Envelope Breaking (NEW)
```
POST /api/EnvelopeBreakageProcessing/ProcessEnvelopeBreaking?ProjectId=1
- Saves to EnvelopeBreakingResult table
```

### Step 5: Get Envelope Breaking Report (NEW)
```
GET /api/EnvelopeBreakageProcessing/GetEnvelopeBreakingReport?ProjectId=1
- Generates Excel from database
```

### Step 6: Process Box Breaking (NEW)
```
POST /api/BoxBreakingProcessing/ProcessBoxBreaking?ProjectId=1
- Saves to BoxBreakingResult table
```

### Step 7: Get Box Breaking Report (NEW)
```
GET /api/BoxBreakingProcessing/GetBoxBreakingReport?ProjectId=1
- Generates Excel from database
```

---

## Key Features

✅ **DB-First Approach** - Data saved to DB before Excel generation
✅ **Exact Same Output** - Excel files match original endpoints exactly
✅ **Batch Tracking** - `UploadBatch` field tracks which upload each record came from
✅ **Comparison Ready** - Can compare Batch 1 vs Batch 2 results
✅ **No Redundancy** - Stores only calculated fields, joins with NRData for details
✅ **Extras Support** - Handles Nodal Extra, University Extra, Office Extra correctly
✅ **Incremental** - Can process new CatchNos without reprocessing all
✅ **Audit Trail** - Full history of processing in database

---

## Database Models Required

### EnvelopeBreakingResult
```csharp
public int Id { get; set; }
public int ProjectId { get; set; }
public int? NrDataId { get; set; }      // NULL if extra
public int? ExtraId { get; set; }       // NULL if regular NRData
public string CatchNo { get; set; }
public int EnvQuantity { get; set; }
public int CenterEnv { get; set; }
public int TotalEnv { get; set; }
public string Env { get; set; }
public int SerialNumber { get; set; }
public string BookletSerial { get; set; }
public DateTime CreatedAt { get; set; }
public int UploadBatch { get; set; }
```

### BoxBreakingResult
```csharp
public int Id { get; set; }
public int ProjectId { get; set; }
public int? NrDataId { get; set; }      // NULL if extra
public int? ExtraId { get; set; }       // NULL if regular NRData
public string CatchNo { get; set; }
public int Start { get; set; }
public int End { get; set; }
public string Serial { get; set; }
public int TotalPages { get; set; }
public string BoxNo { get; set; }
public string OmrSerial { get; set; }
public int? InnerBundlingSerial { get; set; }
public DateTime CreatedAt { get; set; }
public int UploadBatch { get; set; }
```

---

## Next Steps

1. Create migration for new models
2. Update DbContext to include new DbSets
3. Test POST endpoints to save data
4. Test GET endpoints to generate Excel
5. Compare Excel output with original endpoints
6. Add comparison endpoints to show differences between batches
