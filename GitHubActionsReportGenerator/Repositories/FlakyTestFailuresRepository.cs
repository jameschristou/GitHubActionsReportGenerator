using NHibernate;
using NHibernate.Transform;

namespace GitHubActionsReportGenerator.Repositories
{
    public class FlakyTestFailure
    {
        public virtual string TestName { get; set; }
        public virtual long RunId { get; set; }
        public virtual int RunAttempt { get; set; }
        public virtual string JobUrl { get; set; }
        public virtual DateTime StartedAtUtc { get; set; }

    }

	public class FlakyTestFailuresRepository
	{
		private readonly ISessionFactory _sessionFactory;

		public FlakyTestFailuresRepository(ISessionFactory sessionFactory)
		{
			_sessionFactory = sessionFactory;
		}

		public async Task<IEnumerable<FlakyTestFailure>> GetFlakyTestFailures(DateTime endDate)
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
			tests_with_failures AS (
				SELECT	test_results_last_4weeks.RunId,
						test_results_last_4weeks.Name,
						Result
				FROM	test_results_last_4weeks
				JOIN	
				(
					SELECT	RunId,
							Name
					FROM	test_results_last_4weeks
					WHERE	Result = 'Failed'
				) AS tests_that_failed
				ON tests_that_failed.RunId = test_results_last_4weeks.RunId AND tests_that_failed.Name = test_results_last_4weeks.Name
				GROUP BY test_results_last_4weeks.RunId, test_results_last_4weeks.Name, Result
			)

			SELECT	test_results_last_4weeks.Name AS TestName,
					Url AS JobUrl,
					test_results_last_4weeks.RunId,
					RunAttempt,
					StartedAtUtc
			FROM test_results_last_4weeks
			JOIN
			(
				-- find tests that have both failed and succeeded within the same run
				SELECT	TestsThatFailedWithinARun.RunId,
						TestsThatFailedWithinARun.Name
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
			) AS LikelyFlakyTests
			ON LikelyFlakyTests.Name = test_results_last_4weeks.Name
				AND LikelyFlakyTests.RunId = test_results_last_4weeks.RunId
			Where Result = 'Failed'
			ORDER by StartedAtUtc DESC
            ")
			.AddScalar("TestName", NHibernateUtil.String)
			.AddScalar("RunId", NHibernateUtil.Int64)
			.AddScalar("RunAttempt", NHibernateUtil.Int32)
			.AddScalar("JobUrl", NHibernateUtil.String)
			.AddScalar("StartedAtUtc", NHibernateUtil.DateTime)
			.SetParameter("endDate", endDate)
			.SetResultTransformer(Transformers.AliasToBean<FlakyTestFailure>())
			.SetTimeout(180);

			return await query.ListAsync<FlakyTestFailure>();
		}
	}
}
