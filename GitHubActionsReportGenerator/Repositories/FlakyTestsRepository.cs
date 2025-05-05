using NHibernate;
using NHibernate.Transform;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitHubActionsReportGenerator.Repositories
{
    public class FlakyTest
    {
        public virtual string TestName { get; set; }
        public virtual int AvgDurationSeconds { get; set; }
        public virtual int MaxDurationSeconds { get; set; }
        public virtual int MinDurationSeconds { get; set; }
        public virtual int FailureCount { get; set; }
        public virtual int SuccessCount { get; set; }
        public virtual int FailurePercentage { get; set; }
        public virtual int NumberOfRunsImpacted { get; set; }
        public virtual int FailureCountLast4Weeks { get; set; }
    }

    public class FlakyTestsRepository
    {
        private readonly ISessionFactory _sessionFactory;
        public FlakyTestsRepository(ISessionFactory sessionFactory)
        {
            _sessionFactory = sessionFactory;
        }

        public async Task<IEnumerable<FlakyTest>> GetFlakyTests(DateTime endDate)
        {
            using var session = _sessionFactory.OpenSession();

            var query = session.CreateSQLQuery(@"
            DECLARE	@fromDate DATETIME = DATEADD(d, -7, :endDate),
					@fromDate4weeks DATETIME = DATEADD(d, -28, :endDate),
					@endDate DATETIME = :endDate;

			WITH unfiltered_cypress_test_results_last_4weeks AS (
				SELECT	SUBSTRING(wrj.Name, LEN(wrj.Name) - CHARINDEX('/', REVERSE(wrj.Name)) + 3, LEN(wrj.Name)) AS JobName,
						tr.Name AS TestName,
						DurationMs/1000 AS DurationSeconds,
						Result,
						wrj.Url,
						wrj.StartedAtUtc,
						wrj.RunId,
						wrj.RunAttempt
				FROM [GHAData].[dbo].[WorkflowRun] r
				JOIN [GHAData].[dbo].[WorkflowRunJob] wrj
				ON wrj.WorkflowRunId = r.Id
				JOIN [GHAData].[dbo].[TestResult] tr
				ON tr.WorkflowRunJobId = wrj.Id
				WHERE r.Conclusion in ('success', 'failure') 
					AND wrj.StartedAtUtc > @fromDate4weeks
					AND wrj.Name LIKE '%cypress%'
			),
			unfiltered_test_results_last_4weeks AS (
				SELECT	LEFT(JobName, CHARINDEX('-', JobName) - 1) + ' ' + TestName AS Name,
						DurationSeconds,
						Result,
						Url,
						StartedAtUtc,
						RunId,
						RunAttempt
				FROM	unfiltered_cypress_test_results_last_4weeks

				UNION

				-- API tests
				SELECT	tr.Name,
						DurationMs/1000 AS DurationSeconds,
						Result,
						wrj.Url,
						wrj.StartedAtUtc,
						wrj.RunId,
						wrj.RunAttempt
				FROM [GHAData].[dbo].[WorkflowRun] r
				JOIN [GHAData].[dbo].[WorkflowRunJob] wrj
				ON wrj.WorkflowRunId = r.Id
				JOIN [GHAData].[dbo].[TestResult] tr
				ON tr.WorkflowRunJobId = wrj.Id
				WHERE r.Conclusion in ('success', 'failure') 
					AND wrj.StartedAtUtc > @fromDate4weeks
					AND wrj.StartedAtUtc < @endDate
					AND wrj.Name LIKE '% API %'
			),
			test_results_to_ignore AS (
				-- if a run attempt has had more than 10 failed tests, it's likely there's been some other issues and it's not related to flaky tests
				SELECT	RunId,
						RunAttempt
				FROM	unfiltered_test_results_last_4weeks
				WHERE	Result = 'Failed'
				GROUP BY RunId, RunAttempt
				HAVING COUNT(RunAttempt) > 10
			),
			test_results_last_4weeks AS (
				-- remove all the test results we want to ignore
				SELECT	Name,
						DurationSeconds,
						Result,
						Url,
						StartedAtUtc,
						tr.RunId,
						tr.RunAttempt
				FROM	unfiltered_test_results_last_4weeks tr
				LEFT JOIN test_results_to_ignore ignore
				ON ignore.RunId = tr.RunId AND ignore.RunAttempt = tr.RunAttempt
				WHERE	ignore.RunId IS NULL
			),
			test_results_last_week AS (
				SELECT	Name,
						DurationSeconds,
						Result,
						RunId
				FROM	test_results_last_4weeks
				WHERE	StartedAtUtc > @fromDate
						AND StartedAtUtc < @endDate
			),
			tests_with_failures AS (
				SELECT	test_results_last_week.RunId,
						test_results_last_week.Name,
						Result
				FROM	test_results_last_week
				JOIN	
				(
					SELECT	RunId,
							Name
					FROM	test_results_last_week
					WHERE	Result = 'Failed'
				) AS tests_that_failed
				ON tests_that_failed.RunId = test_results_last_week.RunId AND tests_that_failed.Name = test_results_last_week.Name
				GROUP BY test_results_last_week.RunId, test_results_last_week.Name, Result
			)

			SELECT	Durations.Name AS TestName, 
					Durations.AvgDurationSeconds,
					Durations.MaxDurationSeconds,
					Durations.MinDurationSeconds,
					ISNULL(Failures.FailureCount, 0) AS FailureCount, 
					ISNULL(Successes.SuccessCount, 0) AS SuccessCount,
					ISNULL(Failures.FailureCount, 0)*100/(ISNULL(Failures.FailureCount, 0) + ISNULL(Successes.SuccessCount, 0)) AS FailurePercentage,
					ISNULL(RunsImpacted.NumberOfRunsImpacted, 0) AS NumberOfRunsImpacted,
					ISNULL(FailuresLast4Weeks.FailureCount, 0) AS FailureCountLast4Weeks
			FROM
			(
				SELECT	Name,
						AVG(DurationSeconds) AS AvgDurationSeconds,
						MAX(DurationSeconds) AS MaxDurationSeconds,
						MIN(DurationSeconds) AS MinDurationSeconds
				FROM	test_results_last_week
				GROUP BY Name
			) AS Durations
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Result) AS FailureCount
				FROM	test_results_last_week
				WHERE	Result = 'Failed'
				GROUP BY Name, Result
			) AS Failures
			ON Failures.Name = Durations.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Result) AS SuccessCount
				FROM	test_results_last_week
				WHERE	Result = 'Passed'
				GROUP BY Name, Result
			) AS Successes
			ON Successes.Name = Durations.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Name) AS NumberOfRunsImpacted
				FROM (
					SELECT	Name,
							COUNT(RunId) AS NumberOfRunsImpacted
					FROM	test_results_last_week
					WHERE	Result = 'Failed'
					GROUP BY Name, RunId
				) AS RunsImpacted
				GROUP BY Name
			) AS RunsImpacted
			ON RunsImpacted.Name = Durations.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Result) AS FailureCount
				FROM	test_results_last_4weeks
				WHERE	Result = 'Failed'
				GROUP BY Name, Result
			) AS FailuresLast4Weeks
			ON FailuresLast4Weeks.Name = Durations.Name
			JOIN
			(
				-- find tests that have both failed and succeeded within the same run
				SELECT TestsThatFailedWithinARun.Name
				FROM
				(
					SELECT	RunId,
							Name
					FROM	tests_with_failures
					WHERE	Result = 'Passed'
					GROUP BY RunId, Name
				) AS TestsThatPassedWithinARun
				JOIN
				(
					SELECT	RunId,
							Name
					FROM	tests_with_failures
					WHERE	Result = 'Failed'
					GROUP BY RunId, Name -- necessary to eliminate duplicates where test failed and passed in multiple runs
				) AS TestsThatFailedWithinARun
				ON TestsThatFailedWithinARun.RunId = TestsThatPassedWithinARun.RunId
					AND TestsThatFailedWithinARun.Name = TestsThatPassedWithinARun.Name
				GROUP BY TestsThatFailedWithinARun.Name
			) AS LikelyFlakyTests
			ON LikelyFlakyTests.Name = Durations.Name
			ORDER by --MinDurationSeconds desc
					 FailureCount	desc
            ")
            .AddScalar("TestName", NHibernateUtil.String)
            .AddScalar("AvgDurationSeconds", NHibernateUtil.Int32)
            .AddScalar("MaxDurationSeconds", NHibernateUtil.Int32)
            .AddScalar("MinDurationSeconds", NHibernateUtil.Int32)
            .AddScalar("FailureCount", NHibernateUtil.Int32)
            .AddScalar("SuccessCount", NHibernateUtil.Int32)
            .AddScalar("FailurePercentage", NHibernateUtil.Int32)
            .AddScalar("NumberOfRunsImpacted", NHibernateUtil.Int32)
            .AddScalar("FailureCountLast4Weeks", NHibernateUtil.Int32)
            .SetParameter("endDate", endDate)
            .SetResultTransformer(Transformers.AliasToBean<FlakyTest>())
			.SetTimeout(180);

            return await query.ListAsync<FlakyTest>();
        }
    }
}
