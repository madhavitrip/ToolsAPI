# Handling Extras in Box Breaking & Envelope Breaking Results

## The Problem

**ExtraEnvelopes Table Structure:**
```
Id | ProjectId | CatchNo | ExtraId | Quantity | InnerEnvelope | OuterEnvelope
1  | 1         | C001    | 1       | 50       | {...}         | {...}
2  | 1         | C001    | 2       | 30       | {...}         | {...}
3  | 1         | C001    | 3       | 20       | {...}         | {...}
```

- **No NrDataId** - Extras are identified by CatchNo + ExtraId
- **ExtraId meanings:**
  - 1 = Nodal Extra
  - 2 = University Extra
  - 3 = Office Extra
- **CatchNo can repeat** - Same CatchNo can have multiple extras

**In EnvelopeBreakage endpoint:**
- Extras are added to `resultList` with special CenterCode values ("Nodal Extra", "University Extra", "Office Extra")
- They appear as separate rows in the Excel output

**In Replication endpoint:**
- Extras are NOT included (they're already in EnvelopeBreaking.xlsx from previous step)
- Replication reads EnvelopeBreaking.xlsx which already has extras mixed in

---

## Solution: Add ExtraId Field to Result Tables

### Modified EnvelopeBreakingResult

```csharp
public class EnvelopeBreakingResult
{
    public int Id { get; set; }
    public int ProjectId { get; set; }
    
    // Can be either NrDataId OR ExtraId (but not both)
    public int? NrDataId { get; set; }      // NULL if this is an extra
    public int? ExtraId { get; set; }       // NULL if this is regular NRData (1=Nodal, 2=University, 3=Office)
    
    public string CatchNo { get; set; }     // Store this for reference (also in NRData/ExtraEnvelopes)
    
    // Calculated fields
    public int EnvQuantity { get; set; }
    public int CenterEnv { get; set; }
    public int TotalEnv { get; set; }
    public string Env { get; set; }         // "1/2", "2/2"
    public int SerialNumber { get; set; }
    public string BookletSerial { get; set; }
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public int UploadBatch { get; set; }
}
```

### Modified BoxBreakingResult

```csharp
public class BoxBreakingResult
{
    public int Id { get; set; }
    public int ProjectId { get; set; }
    
    // Can be either NrDataId OR ExtraId (but not both)
    public int? NrDataId { get; set; }      // NULL if this is an extra
    public int? ExtraId { get; set; }       // NULL if this is regular NRData (1=Nodal, 2=University, 3=Office)
    
    public string CatchNo { get; set; }     // Store this for reference
    
    // Calculated fields
    public int Start { get; set; }
    public int End { get; set; }
    public string Serial { get; set; }      // "1 to 5"
    public int TotalPages { get; set; }
    public string BoxNo { get; set; }
    public string OmrSerial { get; set; }
    public int? InnerBundlingSerial { get; set; }
    
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public int UploadBatch { get; set; }
}
```

---

## How to Store Data

### In EnvelopeBreakage Endpoint

```csharp
// After generating resultList, save to DB

var envelopeResults = new List<EnvelopeBreakingResult>();

foreach (var item in resultList)
{
    var dict = (Dictionary<string, object>)item;
    
    // Determine if this is an extra or regular NRData
    bool isExtra = dict.ContainsKey("ExtraAttached") && (bool)dict["ExtraAttached"];
    
    int? nrDataId = null;
    int? extraId = null;
    string catchNo = dict["CatchNo"]?.ToString();
    
    if (isExtra)
    {
        // This is an extra row
        extraId = (int)dict["ExtraId"];  // 1, 2, or 3
        // CatchNo is already set
    }
    else
    {
        // This is a regular NRData row
        var nrRow = nrData.FirstOrDefault(n => n.CatchNo == catchNo);
        nrDataId = nrRow?.Id;
    }
    
    envelopeResults.Add(new EnvelopeBreakingResult
    {
        ProjectId = ProjectId,
        NrDataId = nrDataId,
        ExtraId = extraId,
        CatchNo = catchNo,
        EnvQuantity = (int)dict["EnvQuantity"],
        CenterEnv = (int)dict["CenterEnv"],
        TotalEnv = (int)dict["TotalEnv"],
        Env = dict["Env"]?.ToString(),
        SerialNumber = (int)dict["SerialNumber"],
        BookletSerial = dict["BookletSerial"]?.ToString(),
        UploadBatch = currentBatch
    });
}

_context.EnvelopeBreakingResults.AddRange(envelopeResults);
await _context.SaveChangesAsync();
```

### In Replication Endpoint

```csharp
// After generating finalWithBoxes, save to DB

var boxResults = new List<BoxBreakingResult>();
int serial = 1;

foreach (var item in finalWithBoxes)
{
    string catchNo = item.CatchNo;
    
    // Determine if this is an extra or regular NRData
    // Check if CenterCode indicates an extra
    bool isExtra = item.CenterCode?.ToString().Contains("Extra") ?? false;
    
    int? nrDataId = null;
    int? extraId = null;
    
    if (isExtra)
    {
        // Parse ExtraId from CenterCode or find from extras list
        // "Nodal Extra" = 1, "University Extra" = 2, "Office Extra" = 3
        extraId = item.CenterCode switch
        {
            "Nodal Extra" => 1,
            "University Extra" => 2,
            "Office Extra" => 3,
            _ => null
        };
    }
    else
    {
        // Regular NRData
        var nrRow = nrData.FirstOrDefault(n => n.CatchNo == catchNo);
        nrDataId = nrRow?.Id;
    }
    
    boxResults.Add(new BoxBreakingResult
    {
        ProjectId = ProjectId,
        NrDataId = nrDataId,
        ExtraId = extraId,
        CatchNo = catchNo,
        Start = item.Start,
        End = item.End,
        Serial = item.Serial,
        TotalPages = item.TotalPages,
        BoxNo = item.BoxNo?.ToString(),
        OmrSerial = item.OmrSerial,
        InnerBundlingSerial = item.InnerBundlingSerial,
        UploadBatch = currentBatch
    });
}

_context.BoxBreakingResults.AddRange(boxResults);
await _context.SaveChangesAsync();
```

---

## Querying the Data

### Get all results for a CatchNo (including extras)

```csharp
var catchResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.CatchNo == "C001")
    .ToListAsync();

// Results will include:
// - Regular NRData rows (NrDataId = 1, ExtraId = NULL)
// - Nodal Extra rows (NrDataId = NULL, ExtraId = 1)
// - University Extra rows (NrDataId = NULL, ExtraId = 2)
// - Office Extra rows (NrDataId = NULL, ExtraId = 3)
```

### Get only regular NRData (exclude extras)

```csharp
var regularResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.NrDataId != null)
    .ToListAsync();
```

### Get only extras

```csharp
var extraResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.ExtraId != null)
    .ToListAsync();
```

### Get specific extra type (e.g., Nodal Extra)

```csharp
var nodalExtras = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.ExtraId == 1)
    .ToListAsync();
```

### Join with NRData to get full details

```csharp
var fullResults = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.NrDataId != null)
    .Join(_context.NRDatas,
        r => r.NrDataId,
        n => n.Id,
        (r, n) => new
        {
            Result = r,
            NRData = n,
            CatchNo = n.CatchNo,
            NodalCode = n.NodalCode,
            ExamTime = n.ExamTime
        })
    .ToListAsync();
```

### Join with ExtraEnvelopes to get extra details

```csharp
var extraDetails = await _context.EnvelopeBreakingResults
    .Where(r => r.ProjectId == ProjectId && r.ExtraId != null)
    .Join(_context.ExtrasEnvelope,
        r => new { r.ProjectId, r.CatchNo, r.ExtraId },
        e => new { e.ProjectId, e.CatchNo, ExtraId = (int?)e.ExtraId },
        (r, e) => new
        {
            Result = r,
            Extra = e,
            ExtraTypeName = r.ExtraId == 1 ? "Nodal Extra" : 
                           r.ExtraId == 2 ? "University Extra" : "Office Extra"
        })
    .ToListAsync();
```

---

## Database Schema

### EnvelopeBreakingResult Table
```
Id | ProjectId | NrDataId | ExtraId | CatchNo | EnvQuantity | CenterEnv | TotalEnv | Env   | SerialNumber | BookletSerial | UploadBatch | CreatedAt
1  | 1         | 1        | NULL    | C001    | 50          | 1         | 2        | 1/2   | 1            | 1-50          | 1           | 2025-03-10
2  | 1         | 1        | NULL    | C001    | 50          | 2         | 2        | 2/2   | 2            | 51-100        | 1           | 2025-03-10
3  | 1         | NULL     | 1       | C001    | 25          | 1         | 1        | 1/1   | 3            | 101-125       | 1           | 2025-03-10
4  | 1         | NULL     | 2       | C001    | 30          | 1         | 1        | 1/1   | 4            | 126-155       | 1           | 2025-03-10
5  | 1         | 2        | NULL    | C002    | 75          | 1         | 1        | 1/1   | 1            | 156-230       | 1           | 2025-03-10
```

### BoxBreakingResult Table
```
Id | ProjectId | NrDataId | ExtraId | CatchNo | Start | End | Serial    | TotalPages | BoxNo | OmrSerial | InnerBundlingSerial | UploadBatch | CreatedAt
1  | 1         | 1        | NULL    | C001    | 1     | 5   | 1 to 5    | 250        | 1     | 1-50      | 1                   | 1           | 2025-03-10
2  | 1         | 1        | NULL    | C001    | 6     | 10  | 6 to 10   | 250        | 2     | 51-100    | 1                   | 1           | 2025-03-10
3  | 1         | NULL     | 1       | C001    | 11    | 12  | 11 to 12  | 125        | 3     | 101-125   | 1                   | 1           | 2025-03-10
4  | 1         | NULL     | 2       | C001    | 13    | 14  | 13 to 14  | 150        | 4     | 126-155   | 1                   | 1           | 2025-03-10
5  | 1         | 2        | NULL    | C002    | 15    | 19  | 15 to 19  | 375        | 5     | 156-230   | 2                   | 1           | 2025-03-10
```

---

## Key Points

✅ **NrDataId OR ExtraId** - One is set, the other is NULL
✅ **CatchNo stored** - For easy reference and filtering
✅ **No data duplication** - Join with NRData/ExtraEnvelopes when you need full details
✅ **Easy filtering** - Can query regular vs extras separately
✅ **Comparison still works** - Compare UploadBatch 1 vs 2
✅ **Audit trail** - Know which extra type each row is
