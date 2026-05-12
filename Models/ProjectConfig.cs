using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ProjectConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        public List<int> Modules { get; set; }
        public int ProjectId { get; set; }
        public string Envelope { get; set; }
        public List<int> BoxBreakingCriteria { get; set; }
        public List<int> DuplicateRemoveFields { get; set; }
        public List<int> SortingBoxReport { get; set; }
        public List<int> EnvelopeMakingCriteria { get; set; }
        public int BoxCapacity { get; set; }
        public List<int> DuplicateCriteria { get; set; }
        public double Enhancement { get; set; }
        public int BoxNumber { get; set; }
        public int OmrSerialNumber { get; set; }
        public bool ResetOnSymbolChange { get; set; }
        public bool IsInnerBundlingDone { get; set; }
        public List<int>? InnerBundlingCriteria { get; set; }
        public bool ResetOmrSerialOnCatchChange { get; set; }
        public int? BookletSerialNumber { get; set; }
        public bool? ResetBookletSerialOnCatchChange { get; set; }
        public List<int> MssTypes { get; set; } = new List<int>();
        public string MssAttached {  get; set; }
        public bool RoundOffBeforeEnhancement { get; set; }
    }

    public static class PipelineNavigator
    {
        // Define canonical step IDs
        public const int STEP_UPLOADED = 0;
        public const int STEP_DUP_PARTIAL = 1;
        public const int STEP_AWAITING_EXTRA = 2; // Extra Configuration
        public const int STEP_AWAITING_ENV = 3;   // Envelope Setup
        public const int STEP_AWAITING_BOX = 4;   // Box Breaking
        public const int STEP_DONE = 5;

        // IDs assigned in DB Modules table
        public const int MODULE_ID_DUP = 1;
        public const int MODULE_ID_EXTRA = 2;
        public const int MODULE_ID_ENV = 3;
        public const int MODULE_ID_BOX = 4;

        /// <summary>
        /// Determines the first valid step upon data upload.
        /// </summary>
        public static int GetInitialStep(List<int>? activeModules)
        {
            var active = activeModules ?? new List<int>();

            if (active.Contains(MODULE_ID_DUP)) return STEP_UPLOADED;
            if (active.Contains(MODULE_ID_EXTRA)) return STEP_AWAITING_EXTRA;
            if (active.Contains(MODULE_ID_ENV)) return STEP_AWAITING_ENV;
            if (active.Contains(MODULE_ID_BOX)) return STEP_AWAITING_BOX;

            return STEP_DONE;
        }

        /// <summary>
        /// Determines the next valid step after completing a specific step.
        /// </summary>
        public static int GetNextStep(int completedStep, List<int>? activeModules)
        {
            var active = activeModules ?? new List<int>();

            // Duplicate Tool finished
            if (completedStep == STEP_DUP_PARTIAL)
            {
                if (active.Contains(MODULE_ID_EXTRA)) return STEP_AWAITING_EXTRA;
                if (active.Contains(MODULE_ID_ENV)) return STEP_AWAITING_ENV;
                if (active.Contains(MODULE_ID_BOX)) return STEP_AWAITING_BOX;
                return STEP_DONE;
            }

            // Extra Configuration finished
            if (completedStep == STEP_AWAITING_EXTRA)
            {
                if (active.Contains(MODULE_ID_ENV)) return STEP_AWAITING_ENV;
                if (active.Contains(MODULE_ID_BOX)) return STEP_AWAITING_BOX;
                return STEP_DONE;
            }

            // Envelope Setup finished
            if (completedStep == STEP_AWAITING_ENV)
            {
                if (active.Contains(MODULE_ID_BOX)) return STEP_AWAITING_BOX;
                return STEP_DONE;
            }

            return STEP_DONE;
        }

        /// <summary>
        /// Gets an array of all eligible incoming step values for a given target step allowing re-runs/fallbacks.
        /// </summary>
        public static int[] GetEligiblePickupSteps(int targetStep)
        {
            // If a module expects step 2, it should pick up 2.
            // Under some workflows, older data might have skipped straight to 4. We want to be inclusive.
            switch (targetStep)
            {
                case STEP_AWAITING_EXTRA:
                    return new[] { STEP_AWAITING_EXTRA, STEP_AWAITING_ENV, STEP_AWAITING_BOX };
                case STEP_AWAITING_ENV:
                    return new[] { STEP_AWAITING_ENV, STEP_AWAITING_BOX, STEP_AWAITING_EXTRA };
                case STEP_AWAITING_BOX:
                    return new[] { STEP_AWAITING_BOX };
                default:
                    return new[] { targetStep };
            }
        }
    }
}
