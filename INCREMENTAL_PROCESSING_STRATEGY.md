# Incremental Processing Strategy with Status Tracking

## Current Problem

1. **Duplicate Controller** - Deletes duplicate rows instead of marking them
2. **No Status Tracking** - Can't tell which NRData has been processed
3. **All-or-Nothing Processing** - Must reprocess entire project when adding new CatchNo
4. **No Comparison** - Can't compare old vs new processing results

---

## Solution: Add Status Field to NRData

### Modify NRData Model

```csharp
public class NRData
{
    [Key]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    public int Id { get; set; }
    public int ProjectId { get; set; }
    
    // ... existing fields ...
    public string CatchNo { get; set; }
    public string NodalCode { get; set; }
    public string ExamTime { get; set; }
    // ... etc ...
    
    // NEW: Status tracking
    public string ProcessingStatus { get; set; } = "Pending";  // Pending, Processed, Merged, Deleted
    public int? MergedIntoNrDataId { get; set; }  // If merged, which record it merged into
    public DateTime? ProcessedAt { get; set; }
    public int ProcessingBatch { get; set; } = 0;  // Track which upload batch
}
```

---

## How It Works

### Scenario 1: Initial Upload (All CatchNos)

**Step 1: Upload NRData**
- All rows have `ProcessingStatus = "Pending"`, `ProcessingBatch = 1`

**Step 2: Run Duplicate Controller**
- Rows to KEEP: `ProcessingStatus = "Processed"`, `ProcessedAt = now`, `ProcessingBatch = 1`
- Rows to MERGE: `ProcessingStatus = "Merged"`, `MergedIntoNrDataId = [kept row Id]`, `ProcessingBatch = 1`
- ❌ DON'T DELETE - just mark as "Merged"

**Step 3: Run EnvelopeConfiguration**
- Only processes rows where `ProcessingStatus = "Processed"`
- Creates EnvelopeBreakage records

**Step 4: Run EnvelopeBreakage & Replication**
- Only processes rows where `ProcessingStatus = "Processed"`
- Stores results with `ProcessingBatch = 1`

---

### Scenario 2: New CatchNo Added (Incremental)

**Step 1: Upload New NRData**
- New rows have `ProcessingStatus = "Pending"`, `ProcessingBatch = 2`
- Existing rows remain unchanged

**Step 2: Run Duplicate Controller (Incremental)**
```csharp
// Only process NEW rows (Pending status)
var newRows = await _context.NRDatas
    .Where(n => n.ProjectId == ProjectId && n.ProcessingStatus == "Pending")
    .ToListAsync();

// Check for duplicates ONLY within new rows + existing processed rows
var existingProcessed = await _context.NRDatas
    .Where(n => n.ProjectId == ProjectId && n.ProcessingStatus == "Processed")
    .ToListAsync();

// Merge logic...
// Mark new rows as Processed or Merged
```

**Step 3: Run EnvelopeConfiguration (Incremental)**
```csharp
// Only process NEW rows that are now "Processed"
var newProcessedRows = await _context.NRDatas
    .Where(n => n.ProjectId == ProjectId && n.ProcessingStatus == "Processed" && n.ProcessingBatch == 2)
    .ToListAsync();

// Create EnvelopeBreakage only for these new rows
```

**Step 4: Run EnvelopeBreakage & Replication (Incremental)**
```csharp
// Only process NEW rows
var newEnvelopeResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == 2)
    .ToListAsync();

var newBoxResults = await _context.BoxBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == 2)
    .ToListAsync();

// Compare with previous batch
var previousEnvelopeResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == 1)
    .ToListAsync();

// Show differences
```

---

## Benefits

✅ **No Data Loss** - Merged rows are marked, not deleted
✅ **Incremental Processing** - Only process new CatchNos
✅ **Comparison** - Can compare Batch 1 vs Batch 2 results
✅ **Audit Trail** - Know when and what was processed
✅ **Rollback Possible** - Can revert to previous batch if needed
✅ **No Redundancy** - Don't store CatchNo, NodalCode, ExamTime again (they're in NRData)

---

## Database Schema

### NRData Table (Modified)
```
Id | ProjectId | CatchNo | NodalCode | ExamTime | ... | ProcessingStatus | MergedIntoNrDataId | ProcessedAt | ProcessingBatch
1  | 1         | C001    | N1        | 09:00    | ... | Processed        | NULL               | 2025-03-10  | 1
2  | 1         | C001    | N1        | 09:00    | ... | Merged           | 1                  | 2025-03-10  | 1
3  | 1         | C002    | N2        | 10:00    | ... | Processed        | NULL               | 2025-03-10  | 1
4  | 1         | C003    | N3        | 11:00    | ... | Pending          | NULL               | NULL        | 2
```

### EnvelopeBreakingResult Table
```
Id | ProjectId | NrDataId | EnvQuantity | CenterEnv | ... | UploadBatch | CreatedAt
1  | 1         | 1        | 50          | 1         | ... | 1           | 2025-03-10
2  | 1         | 1        | 50          | 2         | ... | 1           | 2025-03-10
3  | 1         | 3        | 75          | 1         | ... | 2           | 2025-03-11
```

### BoxBreakingResult Table
```
Id | ProjectId | NrDataId | Start | End | BoxNo | ... | UploadBatch | CreatedAt
1  | 1         | 1        | 1     | 5   | 1     | ... | 1           | 2025-03-10
2  | 1         | 1        | 6     | 10  | 2     | ... | 1           | 2025-03-10
3  | 1         | 3        | 11    | 15  | 3     | ... | 2           | 2025-03-11
```

---

## Comparison Query Example

```csharp
// Compare results between two batches
var batch1 = await _context.BoxBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == 1)
    .ToListAsync();

var batch2 = await _context.BoxBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.UploadBatch == 2)
    .ToListAsync();

// Find differences
var differences = batch2
    .GroupJoin(batch1, 
        b2 => b2.NrDataId, 
        b1 => b1.NrDataId,
        (b2, b1Group) => new
        {
            NrDataId = b2.NrDataId,
            Batch2BoxNo = b2.BoxNo,
            Batch1BoxNo = b1Group.FirstOrDefault()?.BoxNo,
            Changed = b2.BoxNo != b1Group.FirstOrDefault()?.BoxNo
        })
    .Where(x => x.Changed)
    .ToList();
```

---

## Implementation Steps

1. Add `ProcessingStatus`, `MergedIntoNrDataId`, `ProcessedAt`, `ProcessingBatch` to NRData
2. Create migration
3. Modify DuplicateController to mark as "Merged" instead of deleting
4. Modify EnvelopeConfiguration to only process "Processed" status rows
5. Modify EnvelopeBreakage endpoint to only process "Processed" status rows
6. Modify Replication endpoint to only process "Processed" status rows
7. Add `UploadBatch` to EnvelopeBreakingResult and BoxBreakingResult
8. Create comparison endpoints to show differences between batches
