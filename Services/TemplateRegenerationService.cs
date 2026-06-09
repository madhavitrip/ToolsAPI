using ERPToolsAPI.Data;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Tools.Models;

namespace Tools.Services
{
    /// <summary>
    /// Service to handle template regeneration when imported data fields are edited.
    /// When a field is edited that also exists in an RPT template mapping, 
    /// all templates using that field are marked for regeneration.
    /// </summary>
    public class TemplateRegenerationService
    {
        private readonly ERPToolsDbContext _context;
        private readonly ILoggerService _loggerService;

        public TemplateRegenerationService(ERPToolsDbContext context, ILoggerService loggerService)
        {
            _context = context;
            _loggerService = loggerService;
        }

        /// <summary>
        /// Finds all templates that use the specified field in their RPT mapping.
        /// </summary>
        public async Task<List<int>> FindTemplatesUsingField(int projectId, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return new List<int>();

            try
            {
                var templates = await _context.RPTTemplates
                    .Where(t => 
                        t.ProjectId == projectId && 
                        t.IsActive == true && 
                        (t.IsDeleted == false || t.IsDeleted == null))
                    .ToListAsync();

                var affectedTemplateIds = new List<int>();

                foreach (var template in templates)
                {
                    if (TemplateUsesField(template, fieldName))
                    {
                        affectedTemplateIds.Add(template.TemplateId);
                    }
                }

                return affectedTemplateIds;
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    $"Error finding templates using field '{fieldName}'",
                    ex.Message,
                    nameof(TemplateRegenerationService)
                );
                return new List<int>();
            }
        }

        /// <summary>
        /// Marks specified templates for regeneration by setting ReportStatus to false.
        /// Only marks templates that have already been generated (have reports).
        /// </summary>
        public async Task MarkTemplatesForRegeneration(List<int> templateIds)
        {
            if (!templateIds.Any())
                return;

            try
            {
                // Only mark templates that have already been generated (ReportStatus = true)
                // This ensures we only regenerate existing reports, not create new ones
                var alreadyGeneratedCount = await _context.RPTTemplates
                    .Where(t => templateIds.Contains(t.TemplateId) && t.ReportStatus == true)
                    .CountAsync();

                if (alreadyGeneratedCount == 0)
                {
                    await _loggerService.LogEventAsync(
                        $"No templates with generated reports found for regeneration. Total templates checked: {templateIds.Count}",
                        "RPTTemplate",
                        0,
                        0
                    );
                    return;
                }

                await _context.RPTTemplates
                    .Where(t => templateIds.Contains(t.TemplateId) && t.ReportStatus == true)
                    .ExecuteUpdateAsync(setters => 
                        setters.SetProperty(t => t.ReportStatus, false));

                await _loggerService.LogEventAsync(
                    $"Marked {alreadyGeneratedCount} generated templates for regeneration due to field changes",
                    "RPTTemplate",
                    0,
                    0
                );
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error marking templates for regeneration",
                    ex.Message,
                    nameof(TemplateRegenerationService)
                );
            }
        }

        /// <summary>
        /// Checks if a template's mapping contains the specified field.
        /// </summary>
        private bool TemplateUsesField(RPTTemplate template, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(template.RequiredFieldsJson) && 
                string.IsNullOrWhiteSpace(template.ParsedFieldsJson))
                return false;

            try
            {
                // Check ParsedFieldsJson
                if (!string.IsNullOrWhiteSpace(template.ParsedFieldsJson))
                {
                    var parsed = JArray.Parse(template.ParsedFieldsJson);
                    foreach (var field in parsed)
                    {
                        var fieldNameValue = field["name"]?.ToString() ?? 
                                            field["fieldName"]?.ToString() ?? 
                                            field["Name"]?.ToString() ?? "";
                        
                        if (FieldNamesMatch(fieldNameValue, fieldName))
                            return true;
                    }
                }

                // Check RequiredFieldsJson (fields marked as required in template mapping)
                if (!string.IsNullOrWhiteSpace(template.RequiredFieldsJson))
                {
                    var required = JArray.Parse(template.RequiredFieldsJson);
                    foreach (var field in required)
                    {
                        var fieldNameValue = field["name"]?.ToString() ?? 
                                            field["fieldName"]?.ToString() ?? 
                                            field["Name"]?.ToString() ?? "";
                        if (FieldNamesMatch(fieldNameValue, fieldName))
                            return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"Error checking template {template.TemplateId} for field '{fieldName}': {ex.Message}"
                );
                return false;
            }
        }

        /// <summary>
        /// Compares field names case-insensitively and ignoring whitespace.
        /// </summary>
        private bool FieldNamesMatch(string field1, string field2)
        {
            if (string.IsNullOrWhiteSpace(field1) || string.IsNullOrWhiteSpace(field2))
                return false;

            var normalized1 = NormalizeFieldName(field1);
            var normalized2 = NormalizeFieldName(field2);

            return normalized1 == normalized2;
        }

        /// <summary>
        /// Normalizes field names for comparison.
        /// </summary>
        private string NormalizeFieldName(string fieldName)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                return "";

            return fieldName
                .Replace(" ", "")
                .Replace("-", "")
                .Replace("_", "")
                .ToLowerInvariant();
        }

        /// <summary>
        /// Gets all edited field names from a list of changed fields.
        /// </summary>
        public List<string> GetEditedFieldNames(List<string> changedFields)
        {
            if (changedFields == null || !changedFields.Any())
                return new List<string>();

            // Filter out system fields that don't need template regeneration
            var excludeFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "id",
                "projectid",
                "status",
                "steps",
                "uploadlist",
                "nrdataid",
                "createddate",
                "updateddate"
            };

            return changedFields
                .Where(f => !excludeFields.Contains(f))
                .Distinct()
                .ToList();
        }

        /// <summary>
        /// Gets all dependent module templates that should be regenerated if they have reports.
        /// For example, Box Breaking templates depend on Envelope Breaking being complete.
        /// </summary>
        public async Task<List<int>> GetDependentModuleTemplatesWithReports(
            int projectId, 
            List<int> modulesBeingProcessed)
        {
            if (!modulesBeingProcessed.Any())
                return new List<int>();

            try
            {
                var dependencyMap = new Dictionary<int, List<int>>
                {
                    // Module 4 (Envelope Breaking) → Module 5 (Box Breaking) templates
                    { 4, new List<int> { 5 } },
                    // Add more dependencies as needed
                };

                var dependentModules = new HashSet<int>();
                
                foreach (var moduleId in modulesBeingProcessed)
                {
                    if (dependencyMap.TryGetValue(moduleId, out var dependents))
                    {
                        foreach (var dependent in dependents)
                        {
                            dependentModules.Add(dependent);
                        }
                    }
                }

                if (!dependentModules.Any())
                    return new List<int>();

                // Find templates in dependent modules that have already been generated
                var templatesWithReports = await _context.RPTTemplates
                    .Where(t => 
                        t.ProjectId == projectId &&
                        t.IsActive == true &&
                        (t.IsDeleted == false || t.IsDeleted == null) &&
                        t.ReportStatus == true &&  // Already generated
                        t.ModuleIds != null)
                    .ToListAsync();

                var matchingTemplateIds = new List<int>();

                foreach (var template in templatesWithReports)
                {
                    try
                    {
                        var moduleIds = template.ModuleIds ?? new List<int>();

                        if (moduleIds.Any(mid => dependentModules.Contains(mid)))
                        {
                            matchingTemplateIds.Add(template.TemplateId);
                        }
                    }
                    catch { }
                }

                return matchingTemplateIds;
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error finding dependent module templates",
                    ex.Message,
                    nameof(TemplateRegenerationService)
                );
                return new List<int>();
            }
        }

        /// <summary>
        /// Main entry point for handling field changes and regenerating dependent templates.
        /// This version is called from processing controllers (EnvelopeBreaking, BoxBreaking, etc.)
        /// to regenerate templates from dependent modules if they have already been generated.
        /// </summary>
        public async Task HandleProcessingCompleteAndRegenerateDependent(
            int projectId,
            int completedModuleId)
        {
            try
            {
                // Get dependent templates that have already been generated
                var dependentTemplateIds = await GetDependentModuleTemplatesWithReports(
                    projectId,
                    new List<int> { completedModuleId }
                );

                if (dependentTemplateIds.Any())
                {
                    // Mark only already-generated templates for regeneration
                    await _context.RPTTemplates
                        .Where(t => dependentTemplateIds.Contains(t.TemplateId))
                        .ExecuteUpdateAsync(setters =>
                            setters.SetProperty(t => t.ReportStatus, false));

                    await _loggerService.LogEventAsync(
                        $"Marked {dependentTemplateIds.Count} dependent templates for regeneration after module {completedModuleId} completed",
                        "RPTTemplate",
                        0,
                        projectId
                    );
                }
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error in HandleProcessingCompleteAndRegenerateDependent",
                    ex.Message,
                    nameof(TemplateRegenerationService)
                );
            }
        }

        /// <summary>
        /// Main method to handle field changes and trigger template regeneration.
        /// Only regenerates templates that have already been generated (have reports).
        /// Call this after updating NRData fields.
        /// </summary>
        public async Task HandleFieldChangeAndRegenerateTemplates(
            int projectId, 
            List<string> changedFields)
        {
            if (changedFields == null || !changedFields.Any())
                return;

            try
            {
                var editedFields = GetEditedFieldNames(changedFields);

                if (!editedFields.Any())
                    return;

                var allAffectedTemplates = new HashSet<int>();

                // Find templates using any of the changed fields
                foreach (var fieldName in editedFields)
                {
                    var templatesUsingField = await FindTemplatesUsingField(projectId, fieldName);
                    foreach (var templateId in templatesUsingField)
                    {
                        allAffectedTemplates.Add(templateId);
                    }
                }

                // Expand affected templates to include dependent-module templates (e.g., Box Breaking depends on Envelope)
                if (allAffectedTemplates.Any())
                {
                    try
                    {
                        // Determine modules for affected templates
                        var templates = await _context.RPTTemplates
                            .Where(t => allAffectedTemplates.Contains(t.TemplateId))
                            .ToListAsync();

                        var modulesBeingProcessed = new List<int>();
                        foreach (var t in templates)
                        {
                            if (t.ModuleIds != null)
                            {
                                foreach (var mid in t.ModuleIds)
                                {
                                    if (!modulesBeingProcessed.Contains(mid))
                                        modulesBeingProcessed.Add(mid);
                                }
                            }
                        }

                        // Find dependent templates in modules that depend on the affected modules
                        var dependentTemplateIds = await GetDependentModuleTemplatesWithReports(projectId, modulesBeingProcessed);
                        foreach (var dt in dependentTemplateIds)
                            allAffectedTemplates.Add(dt);

                        // Mark all affected templates (including dependents) for regeneration
                        await MarkTemplatesForRegeneration(allAffectedTemplates.ToList());

                        await _loggerService.LogEventAsync(
                            $"Triggered regeneration check for {allAffectedTemplates.Count} templates (including dependents) due to field changes: {string.Join(", ", editedFields)}",
                            "TemplateRegeneration",
                            0,
                            projectId
                        );
                    }
                    catch (Exception ex)
                    {
                        await _loggerService.LogErrorAsync("Error expanding dependent templates for regeneration", ex.Message, nameof(TemplateRegenerationService));
                    }
                }
            }
            catch (Exception ex)
            {
                await _loggerService.LogErrorAsync(
                    "Error in HandleFieldChangeAndRegenerateTemplates",
                    ex.Message,
                    nameof(TemplateRegenerationService)
                );
            }
        }
    }
}
