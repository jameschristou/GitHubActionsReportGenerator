using FluentNHibernate.Mapping;
using NHibernate;
using NHibernate.Criterion;
using NHibernate.Transform;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GitHubActionsReportGenerator.Repositories
{
    public class WeeklyRunSummary
    {
        public virtual DateTime WeekStarting { get; set; }
        public virtual int AvgDurationForSuccessfulRuns { get; set; }
        public virtual int MaxDurationForSuccessfulRuns { get; set; }
        public virtual int MinDurationForSuccessfulRuns { get; set; }
        public virtual int AvgAttemptsForSuccessfulRuns { get; set; }
        public virtual int FailureCount { get; set; }
        public virtual int SuccessCount { get; set; }
        public virtual int SuccessPercentage { get; set; }
        public virtual int FlakyTestCount { get; set; }
        public virtual int MinutesWastedDueToFailingTests { get; set; }
        public virtual int TotalMinutesWasted { get; set; }
    }

	public interface IRunSummaryRepository
	{
		Task<IEnumerable<WeeklyRunSummary>> GetRunSummary(DateTime endDate);

    }

    public class RunSummaryRepository : IRunSummaryRepository
    {
        private readonly ISessionFactory _sessionFactory;

        public RunSummaryRepository(ISessionFactory sessionFactory)
        {
            _sessionFactory = sessionFactory;
        }

        public async Task<IEnumerable<WeeklyRunSummary>> GetRunSummary(DateTime endDate)
        {
            using var session = _sessionFactory.OpenSession();
			var query = session.CreateSQLQuery(@"
                DECLARE @numDaysInDateRange int = 7;

WITH runs AS (
	SELECT	Id,
			RunId,
			Title,
			Url AS RunUrl,
			StartedAtUtc,
			CompletedAtUtc,
			DATEDIFF(minute, StartedAtUtc, CompletedAtUtc) AS DurationMinutes,
			NumAttempts,
			Conclusion,
			DATEADD(day, - (DATEDIFF(hour, StartedAtUtc, :endDate)/(@numDaysInDateRange*24))*@numDaysInDateRange, :endDate) AS DateGroup
	FROM	[GHAData].[dbo].[WorkflowRun] WITH (NOLOCK)
	WHERE	Conclusion in ('success', 'failure') AND StartedAtUtc < :endDate
),

runAttemptsWithFailedTests AS (
	SELECT	r.DateGroup,
			r.Id,
			rj.RunAttempt,
			DATEDIFF(minute, MIN(rj.StartedAtUtc), MAX(rj.CompletedAtUtc)) AS DurationMinutes
	FROM	runs r
	JOIN	[GHAData].[dbo].[WorkflowRunJob] rj WITH (NOLOCK)
		ON	rj.WorkflowRunId = r.Id
	JOIN	[GHAData].[dbo].[TestResult] tr WITH (NOLOCK)
		ON	tr.WorkflowRunJobId = rj.Id
	WHERE	r.Conclusion in ('success', 'failure') AND (rj.Name LIKE 'Staging%' OR rj.Name LIKE 'Integration%') AND tr.Result = 'Failed'
	GROUP BY r.DateGroup, r.Id, rj.RunAttempt
),

runAttemptsWithFailedJobs AS (
	SELECT	r.DateGroup,
			r.Id,
			rj.RunAttempt,
			DATEDIFF(minute, MIN(rj.StartedAtUtc), MAX(rj.CompletedAtUtc)) AS DurationMinutes
	FROM	runs r
	JOIN	[GHAData].[dbo].[WorkflowRunJob] rj WITH (NOLOCK)
		ON	rj.WorkflowRunId = r.Id
	WHERE	r.Conclusion in ('success', 'failure') AND rj.Name NOT LIKE 'Regression1%' AND rj.Conclusion = 'failure'
	GROUP BY r.DateGroup, r.Id, rj.RunAttempt
),

flakyTests AS (
	SELECT	r.DateGroup,
			tr.Name AS TestName
	FROM	runs r
	JOIN	[GHAData].[dbo].[WorkflowRunJob] rj WITH (NOLOCK)
		ON	rj.WorkflowRunId = r.Id
	JOIN	[GHAData].[dbo].[TestResult] tr WITH (NOLOCK)
		ON	tr.WorkflowRunJobId = rj.Id
	WHERE	r.Conclusion in ('success', 'failure') AND (rj.Name LIKE 'Staging%' OR rj.Name LIKE 'Integration%') AND tr.Result = 'Failed'
			AND r.RunId NOT IN (10804295680, 11906090216) -- runId ignore list in case something went wrong with these runs
			AND NOT (rj.RunId = 12207886958 AND RunAttempt = 1)
			AND NOT (rj.RunId = 12138078758 AND RunAttempt = 1)
			AND NOT (rj.RunId = 12160544449 AND RunAttempt = 1)
			AND NOT (rj.RunId = 12178063588 AND RunAttempt = 1)
			AND NOT (rj.RunId = 12943173927 AND RunAttempt = 1)
			AND NOT (rj.RunId = 13052956675)
			AND NOT (rj.RunId = 13051378308 AND RunAttempt = 1)
			AND NOT (rj.RunId = 13008890783 AND RunAttempt = 1)
			AND NOT (rj.RunId = 13245977686 AND RunAttempt = 1)
			AND NOT (rj.RunId = 13331249031 AND RunAttempt = 1)
	GROUP BY r.DateGroup, tr.Name
)

-- we need to be able to group by date ranges. so each run needs to be put into a date range.
-- what if we do a day diff between the report end date and the date of the run? We divide the
-- number of days by 7 and then get the nearest value. So for example if the day difference is 13
-- 13/7 = 1.85. Which means this becomes part of the group with range 2 weeks ago to 1 week ago.
-- We strip out the decimal and that becomes our group. We can convert that to a date by doing that value * 7 and subtracting from the report date

SELECT	DATEADD(dd, DATEDIFF(dd, 0, DATEADD(hh, 10, Durations.DateGroup)) - 7, 0) AS WeekStarting, -- convert from UTC to AU time, then strip out everything other than the date
		AvgDurationForSuccessfulRuns,
		MaxDurationForSuccessfulRuns,
		MinDurationForSuccessfulRuns,
		AvgAttemptsForSuccessfulRuns,
		ISNULL(Failures.FailureCount, 0) AS FailureCount, 
		ISNULL(Successes.SuccessCount, 0) AS SuccessCount, 
		ISNULL(Successes.SuccessCount, 0)*100/(ISNULL(Failures.FailureCount, 0) + ISNULL(Successes.SuccessCount, 0)) AS SuccessPercentage,
		--ISNULL(RunAttemptsWithFailedTests.RunAttemptsWithFailedTests, 0) AS RunAttemptsWithFailedTests,
		ISNULL(FlakyTests.FlakyTestCount, 0) AS FlakyTestCount,
		ISNULL(TotalFailedTestDuration.DurationMinutes, 0) AS MinutesWastedDueToFailingTests,
		ISNULL(TotalFailedDuration.DurationMinutes, 0) AS TotalMinutesWasted
FROM
(	SELECT	DateGroup, 
		AVG(DurationMinutes) AS AvgDurationForSuccessfulRuns,
		MAX(DurationMinutes) AS MaxDurationForSuccessfulRuns,
		MIN(DurationMinutes) AS MinDurationForSuccessfulRuns
	FROM	runs
	WHERE Conclusion = 'success'
	GROUP BY DateGroup
) AS Durations
LEFT JOIN
(	SELECT	DateGroup, 
		ROUND(AVG(CAST(NumAttempts AS FLOAT)), 2) AS AvgAttemptsForSuccessfulRuns
	FROM	runs
	WHERE Conclusion = 'success'
	GROUP BY DateGroup
) AS NumAttempts
ON NumAttempts.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			COUNT(Conclusion) AS FailureCount
	FROM	runs
	WHERE	Conclusion = 'failure'
	GROUP BY DateGroup, Conclusion
) AS Failures
ON Failures.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			COUNT(Conclusion) AS SuccessCount
	FROM	runs
	WHERE	Conclusion = 'success'
	GROUP BY DateGroup, Conclusion
) AS Successes
ON Successes.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			COUNT(DateGroup) AS RunAttemptsWithFailedTests
	FROM	runAttemptsWithFailedTests
	GROUP BY DateGroup
) AS RunAttemptsWithFailedTests
ON RunAttemptsWithFailedTests.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			SUM(DurationMinutes) AS DurationMinutes
	FROM	runAttemptsWithFailedTests
	GROUP BY DateGroup
) AS TotalFailedTestDuration
ON TotalFailedTestDuration.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			SUM(DurationMinutes) AS DurationMinutes
	FROM	runAttemptsWithFailedJobs
	GROUP BY DateGroup
) AS TotalFailedDuration
ON TotalFailedDuration.DateGroup = Durations.DateGroup
LEFT JOIN
(
	SELECT	DateGroup,
			COUNT(DateGroup) AS FlakyTestCount
	FROM	flakyTests
	GROUP BY DateGroup
) AS FlakyTests
ON FlakyTests.DateGroup = Durations.DateGroup
ORDER BY Durations.DateGroup DESC")
            .AddScalar("WeekStarting", NHibernateUtil.DateTime)
			.AddScalar("AvgDurationForSuccessfulRuns", NHibernateUtil.Int32)
			.AddScalar("MaxDurationForSuccessfulRuns", NHibernateUtil.Int32)
			.AddScalar("MinDurationForSuccessfulRuns", NHibernateUtil.Int32)
			.AddScalar("AvgAttemptsForSuccessfulRuns", NHibernateUtil.Int32)
			.AddScalar("FailureCount", NHibernateUtil.Int32)
			.AddScalar("SuccessCount", NHibernateUtil.Int32)
			.AddScalar("SuccessPercentage", NHibernateUtil.Int32)
			.AddScalar("FlakyTestCount", NHibernateUtil.Int32)
			.AddScalar("MinutesWastedDueToFailingTests", NHibernateUtil.Int32)
			.AddScalar("TotalMinutesWasted", NHibernateUtil.Int32)
            .SetParameter("endDate", endDate)
			.SetResultTransformer(Transformers.AliasToBean<WeeklyRunSummary>())
            .SetTimeout(180);

            return await query.ListAsync<WeeklyRunSummary>();
        }
    }
}
