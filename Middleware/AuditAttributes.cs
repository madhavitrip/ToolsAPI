using System;

namespace Tools.Middleware
{
    /// <summary>
    /// Suppresses automatic centralized audit logging on an action or controller.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
    public class NoAuditLogAttribute : Attribute
    {
    }

    /// <summary>
    /// Overrides the default audit log category (which defaults to the controller name).
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = true)]
    public class AuditCategoryAttribute : Attribute
    {
        public string Category { get; }

        public AuditCategoryAttribute(string category)
        {
            Category = category ?? string.Empty;
        }
    }
}
