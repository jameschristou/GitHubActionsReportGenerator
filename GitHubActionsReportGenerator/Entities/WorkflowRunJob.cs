﻿namespace GitHubActionsReportGenerator.Entities
{
    public class WorkflowRunJob
    {
        public virtual int Id { get; set; }
        public virtual long RunId { get; set; }
        public virtual long JobId { get; set; }
        public virtual int RunAttempt { get; set; }
        public virtual string Name { get; set; }
        public virtual string Conclusion { get; set; }
        public virtual DateTime StartedAtUtc { get; set; }
        public virtual DateTime CompletedAtUtc { get; set; }
        public virtual string Url { get; set; }
        public virtual IList<TestResult> TestResults { get; set; }
        public virtual WorkflowRun WorkflowRun { get; set; }
    }
}
