using NHibernate;
using NHibernate.Transform;

namespace GitHubActionsReportGenerator.Repositories
{
    public class SlowTest
    {
        public virtual string TestName { get; set; }
        public virtual int AvgDurationSecondsLastWeek { get; set; }
        public virtual int MaxDurationSecondsLastWeek { get; set; }
        public virtual int MinDurationSecondsLastWeek { get; set; }
        public virtual int AvgDurationSecondsLast4Weeks { get; set; }
        public virtual int MaxDurationSecondsLast4Weeks { get; set; }
        public virtual int MinDurationSecondsLast4Weeks { get; set; }
    }

    public class SlowestTestsRepository
    {
        private readonly ISessionFactory _sessionFactory;

        public SlowestTestsRepository(ISessionFactory sessionFactory)
        {
            _sessionFactory = sessionFactory;
        }

        public async Task<IEnumerable<SlowTest>> GetSlowTests(DateTime endDate)
        {
            using var session = _sessionFactory.OpenSession();

            var query = session.CreateSQLQuery(@"
				DECLARE	@fromDate DATETIME = DATEADD(d, -7, :endDate),
						@fromDate4weeks DATETIME = DATEADD(d, -28, :endDate),
						@endDate DATETIME = :endDate;

				WITH test_results_last_4weeks AS (
					SELECT	tr.Name,
						DurationMs/1000 AS DurationSeconds,
						Result,
						wrj.RunId,
						wrj.StartedAtUtc
					FROM [GHAData].[dbo].[WorkflowRun] r
					JOIN [GHAData].[dbo].[WorkflowRunJob] wrj
					ON wrj.WorkflowRunId = r.Id
					JOIN [GHAData].[dbo].[TestResult] tr
					ON tr.WorkflowRunJobId = wrj.Id
					WHERE r.Conclusion in ('success', 'failure')
						AND tr.Result = 'Passed'
						AND wrj.StartedAtUtc > @fromDate4weeks 
						AND (wrj.Name LIKE '%API%' OR wrj.Name LIKE '%cypress%')
				),
				test_results_last_week AS (
					SELECT	Name,
							DurationSeconds,
							Result,
							RunId
					FROM	test_results_last_4weeks
					WHERE	StartedAtUtc > @fromDate
				)

				SELECT	TOP 20
						DurationsLastWeek.Name AS TestName, 
						DurationsLastWeek.AvgDurationSeconds AS AvgDurationSecondsLastWeek,
						DurationsLastWeek.MaxDurationSeconds AS MaxDurationSecondsLastWeek,
						DurationsLastWeek.MinDurationSeconds AS MinDurationSecondsLastWeek,
						DurationsLast4Weeks.AvgDurationSeconds AS AvgDurationSecondsLast4Weeks,
						DurationsLast4Weeks.MaxDurationSeconds AS MaxDurationSecondsLast4Weeks,
						DurationsLast4Weeks.MinDurationSeconds AS MinDurationSecondsLast4Weeks
				FROM
				(
					SELECT	Name,
							AVG(DurationSeconds) AS AvgDurationSeconds,
							MAX(DurationSeconds) AS MaxDurationSeconds,
							MIN(DurationSeconds) AS MinDurationSeconds
					FROM test_results_last_week
					GROUP BY Name
				) AS DurationsLastWeek
				JOIN
				(
					SELECT	Name,
							AVG(DurationSeconds) AS AvgDurationSeconds,
							MAX(DurationSeconds) AS MaxDurationSeconds,
							MIN(DurationSeconds) AS MinDurationSeconds
					FROM test_results_last_4weeks
					GROUP BY Name
				) AS DurationsLast4Weeks
				ON DurationsLast4Weeks.Name = DurationsLastWeek.Name
				ORDER by MinDurationSecondsLastWeek desc")
            .AddScalar("TestName", NHibernateUtil.String)
            .AddScalar("AvgDurationSecondsLastWeek", NHibernateUtil.Int32)
            .AddScalar("MaxDurationSecondsLastWeek", NHibernateUtil.Int32)
            .AddScalar("MinDurationSecondsLastWeek", NHibernateUtil.Int32)
            .AddScalar("AvgDurationSecondsLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("MaxDurationSecondsLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("MinDurationSecondsLast4Weeks", NHibernateUtil.Int32)
            .SetParameter("endDate", endDate)
            .SetResultTransformer(Transformers.AliasToBean<SlowTest>())
            .SetTimeout(180);

            return await query.ListAsync<SlowTest>();
        }
    }
}
