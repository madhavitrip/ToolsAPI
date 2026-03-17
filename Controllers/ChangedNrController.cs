using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;
using Tools.Models;
using Tools.Services;

namespace Tools.Controllers
{


    [Route("api/[controller]")]
    [ApiController]
    public class ChangedNrController : ControllerBase
    {

        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;

        public ChangedNrController(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        [HttpPost]
        public async Task<IActionResult> PostChangedNRData([FromBody] JsonElement inputData)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                int projectId = inputData.GetProperty("projectId").GetInt32();
                var dataArray = inputData.GetProperty("data").EnumerateArray();

                var extraConfigs = await _context.ExtraConfigurations
                    .Where(x => x.ProjectId == projectId)
                    .ToListAsync();

                var changedNRDatasToAdd = new List<ChangedNRData>();
                var extraEnvelopesToAdd = new List<ExtraEnvelopes>();

                var dataType = typeof(ChangedNRData);

                var properties = dataType
                    .GetProperties()
                    .ToDictionary(p => p.Name.ToLower(), p => p);

                foreach (var item in dataArray)
                {
                    var changedNRData = new ChangedNRData
                    {
                        ProjectId = projectId
                    };

                    var extraData = new Dictionary<string, string>();

                    foreach (var prop in item.EnumerateObject())
                    {
                        string key = prop.Name.Replace(" ", "").ToLower();
                        string value = prop.Value.ToString();

                        if (properties.TryGetValue(key, out var propInfo))
                        {
                            var targetType = Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                            object convertedValue = string.IsNullOrWhiteSpace(value)
                                ? null
                                : Convert.ChangeType(value, targetType);

                            propInfo.SetValue(changedNRData, convertedValue);
                        }
                        else
                        {
                            extraData[prop.Name] = value;
                        }
                    }

                    if (extraData.Any())
                        changedNRData.NRDatas = JsonSerializer.Serialize(extraData);

                    // EXTRA CENTER LOGIC
                    int? extraTypeId = changedNRData.CenterCode switch
                    {
                        "Nodal Extra" => 1,
                        "University Extra" => 2,
                        "Office Extra" => 3,
                        _ => null
                    };

                    if (extraTypeId.HasValue)
                    {
                        var config = extraConfigs.FirstOrDefault(x => x.ExtraType == extraTypeId);

                        if (config != null)
                        {
                            EnvelopeType envelopeType = null;

                            if (!string.IsNullOrWhiteSpace(config.EnvelopeType))
                            {
                                try
                                {
                                    envelopeType = JsonSerializer.Deserialize<EnvelopeType>(config.EnvelopeType);
                                }
                                catch { }
                            }

                            int? innerCapacity = envelopeType != null ? GetEnvelopeCapacity(envelopeType.Inner) : null;
                            int? outerCapacity = envelopeType != null ? GetEnvelopeCapacity(envelopeType.Outer) : null;

                            int roundedQty = changedNRData.Quantity;

                            if (innerCapacity > 0)
                                roundedQty = (int)Math.Ceiling((double)changedNRData.Quantity / innerCapacity.Value) * innerCapacity.Value;
                            else if (outerCapacity > 0)
                                roundedQty = (int)Math.Ceiling((double)changedNRData.Quantity / outerCapacity.Value) * outerCapacity.Value;

                            string innerEnvelope = innerCapacity > 0
                                ? Math.Ceiling((double)roundedQty / innerCapacity.Value).ToString()
                                : null;

                            string outerEnvelope = outerCapacity > 0
                                ? Math.Ceiling((double)roundedQty / outerCapacity.Value).ToString()
                                : null;

                            extraEnvelopesToAdd.Add(new ExtraEnvelopes
                            {
                                ProjectId = projectId,
                                CatchNo = changedNRData.CatchNo,
                                ExtraId = extraTypeId.Value,
                                Quantity = roundedQty,
                                InnerEnvelope = innerEnvelope,
                                OuterEnvelope = outerEnvelope
                            });
                        }

                        continue;
                    }

                    changedNRDatasToAdd.Add(changedNRData);
                }

                if (changedNRDatasToAdd.Any())
                    await _context.ChangedNRData.AddRangeAsync(changedNRDatasToAdd);

                if (extraEnvelopesToAdd.Any())
                    await _context.ExtrasEnvelope.AddRangeAsync(extraEnvelopesToAdd);

                await _context.SaveChangesAsync();

                var triggeredBy = LogHelper.GetTriggeredBy(User);
                _loggerService.LogEvent(
                    "Changed NRData inserted successfully",
                    "ChangedNRData",
                    triggeredBy,
                    projectId,
                    string.Empty,
                    inputData.GetRawText()
                );

                return Ok(new
                {
                    message = "Changed NRData inserted successfully",
                    Count = changedNRDatasToAdd.Count
                });
            }
            catch (Exception ex)
            {
                _loggerService.LogError("Error inserting Changed NRData", ex.Message, nameof(ChangedNrController));
                return StatusCode(500, ex.Message);
            }
        }

        public class EnvelopeType
        {
            public string Inner { get; set; }
            public string Outer { get; set; }
        }

        private int GetEnvelopeCapacity(string envelopeCode)
        {
            if (string.IsNullOrWhiteSpace(envelopeCode))
                return 1; // default to 1 if null or invalid

            // Expecting format like "E10", "E25", etc.
            var numberPart = new string(envelopeCode.Where(char.IsDigit).ToArray());

            return int.TryParse(numberPart, out var capacity) ? capacity : 1;
        }

     


    }
}
