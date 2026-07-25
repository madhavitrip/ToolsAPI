using ERPToolsAPI.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System;
using Tools.Models;

[ApiController]
[Route("api/[controller]")]
public class ProcessStepsController : ControllerBase
{
    private readonly ERPToolsDbContext _context;

    public ProcessStepsController(ERPToolsDbContext context)
    {
        _context = context;
    }

    [HttpPost]
    public async Task<IActionResult> AddProcessStep([FromBody] ProcessSteps model)
    {
        if (!ModelState.IsValid)
        {
            return BadRequest(ModelState);
        }

        if (string.IsNullOrWhiteSpace(model.ProcessName))
        {
            return BadRequest(new
            {
                success = false,
                message = "Process Name is required."
            });
        }

        if (model.Steps <= 0)
        {
            return BadRequest(new
            {
                success = false,
                message = "Steps must be greater than 0."
            });
        }

        var processStep = new ProcessSteps
        {
            ProcessName = model.ProcessName.Trim(),
            Steps = model.Steps
        };

        _context.ProcessSteps.Add(processStep);
        await _context.SaveChangesAsync();

        return Ok(new
        {
            success = true,
            message = "Process step added successfully.",
            data = processStep
        });
    }

    [HttpGet]
    public async Task<IActionResult> GetAllProcessSteps()
    {
        var processSteps = await _context.ProcessSteps
            .OrderBy(x => x.ProcessId)
            .ToListAsync();

        return Ok(new
        {
            success = true,
            message = "Process steps fetched successfully.",
            count = processSteps.Count,
            data = processSteps
        });
    }
}