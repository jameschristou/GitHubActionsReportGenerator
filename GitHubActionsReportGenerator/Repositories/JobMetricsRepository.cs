using NHibernate;
using NHibernate.Transform;

namespace GitHubActionsReportGenerator.Repositories
{
    public class JobMetric
    {
        public virtual string Name { get; set; }
        public virtual int AvgDurationMinutesLastWeek { get; set; }
        public virtual int MaxDurationMinutesLastWeek { get; set; }
        public virtual int MinDurationMinutesLastWeek { get; set; }
        public virtual int AvgDurationMinutesLast4Weeks { get; set; }
        public virtual int MaxDurationMinutesLast4Weeks { get; set; }
        public virtual int MinDurationMinutesLast4Weeks { get; set; }
        public virtual int FailureCountLastWeek { get; set; }
        public virtual int SuccessCountLastWeek { get; set; }
        public virtual int FailurePercentageLastWeek { get; set; }
        public virtual int FailureCountLast4Weeks { get; set; }
        public virtual int SuccessCountLast4Weeks { get; set; }
        public virtual int FailurePercentageLast4Weeks { get; set; }
    }

    public class JobMetricsRepository
    {
        private readonly ISessionFactory _sessionFactory;

        public JobMetricsRepository(ISessionFactory sessionFactory)
        {
            _sessionFactory = sessionFactory;
        }

        public async Task<IEnumerable<JobMetric>> GetJobMetrics(DateTime endDate)
        {
            using var session = _sessionFactory.OpenSession();

            var query = session.CreateSQLQuery(@"
            DECLARE	@fromDate DATETIME = DATEADD(d, -7, :endDate),
					@fromDate4weeks DATETIME = DATEADD(d, -28, :endDate);

			WITH jobsLast4Weeks AS (
			SELECT	Name,
					DATEDIFF(minute, StartedAtUtc, CompletedAtUtc) AS DurationMinutes,
					Conclusion,
					StartedAtUtc
				FROM [GHAData].[dbo].[WorkflowRunJob]
				WHERE StartedAtUtc > @fromDate4weeks
			),
			jobsLastWeek AS (
			SELECT	Name,
					DurationMinutes,
					Conclusion,
					StartedAtUtc
				FROM jobsLast4Weeks
				WHERE StartedAtUtc > @fromDate
			)

			SELECT	DurationsLastWeek.Name, 
					DurationsLastWeek.AvgDurationMinutesLastWeek,
					DurationsLastWeek.MaxDurationMinutesLastWeek, 
					DurationsLastWeek.MinDurationMinutesLastWeek,
					DurationsLast4Weeks.AvgDurationMinutesLast4Weeks,
					DurationsLast4Weeks.MaxDurationMinutesLast4Weeks, 
					DurationsLast4Weeks.MinDurationMinutesLast4Weeks, 
					ISNULL(FailuresLastWeek.FailureCount, 0) AS FailureCountLastWeek, 
					ISNULL(SuccessesLastWeek.SuccessCount, 0) AS SuccessCountLastWeek, 
					ISNULL(FailuresLastWeek.FailureCount, 0)*100/(ISNULL(FailuresLastWeek.FailureCount, 0) + ISNULL(SuccessesLastWeek.SuccessCount, 0)) AS FailurePercentageLastWeek,
					ISNULL(FailuresLast4Weeks.FailureCount, 0) AS FailureCountLast4Weeks, 
					ISNULL(SuccessesLast4Weeks.SuccessCount, 0) AS SuccessCountLast4Weeks, 
					ISNULL(FailuresLast4Weeks.FailureCount, 0)*100/(ISNULL(FailuresLast4Weeks.FailureCount, 0) + ISNULL(SuccessesLast4Weeks.SuccessCount, 0)) AS FailurePercentageLast4Weeks
			FROM
			(
				SELECT	Name, 
						AVG(DurationMinutes) AS AvgDurationMinutesLastWeek,
						MAX(DurationMinutes) AS MaxDurationMinutesLastWeek,
						MIN(DurationMinutes) AS MinDurationMinutesLastWeek
				FROM jobsLastWeek
				WHERE conclusion = 'Success'
				GROUP BY Name
			) AS DurationsLastWeek
			LEFT JOIN
			(
				SELECT	Name, 
						AVG(DurationMinutes) AS AvgDurationMinutesLast4Weeks,
						MAX(DurationMinutes) AS MaxDurationMinutesLast4Weeks,
						MIN(DurationMinutes) AS MinDurationMinutesLast4Weeks
				FROM jobsLast4Weeks
				WHERE conclusion = 'Success'
				GROUP BY Name
			) AS DurationsLast4Weeks
			ON DurationsLast4Weeks.Name = DurationsLastWeek.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Conclusion) AS FailureCount
				FROM	jobsLastWeek
				WHERE	Conclusion = 'Failure'
				GROUP BY Name, Conclusion
			) AS FailuresLastWeek
			ON FailuresLastWeek.Name = DurationsLastWeek.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Conclusion) AS SuccessCount
				FROM	jobsLastWeek
				WHERE	Conclusion = 'Success'
				GROUP BY Name, Conclusion
			) AS SuccessesLastWeek
			ON SuccessesLastWeek.Name = DurationsLastWeek.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Conclusion) AS FailureCount
				FROM	jobsLast4Weeks
				WHERE	Conclusion = 'Failure'
				GROUP BY Name, Conclusion
			) AS FailuresLast4Weeks
			ON FailuresLast4Weeks.Name = DurationsLastWeek.Name
			LEFT JOIN
			(
				SELECT	Name,
						COUNT(Conclusion) AS SuccessCount
				FROM	jobsLast4Weeks
				WHERE	Conclusion = 'Success'
				GROUP BY Name, Conclusion
			) AS SuccessesLast4Weeks
			ON SuccessesLast4Weeks.Name = DurationsLastWeek.Name
			ORDER BY FailurePercentageLastWeek desc
            ")
            .AddScalar("Name", NHibernateUtil.String)
            .AddScalar("AvgDurationMinutesLastWeek", NHibernateUtil.Int32)
            .AddScalar("MaxDurationMinutesLastWeek", NHibernateUtil.Int32)
            .AddScalar("MinDurationMinutesLastWeek", NHibernateUtil.Int32)
			.AddScalar("AvgDurationMinutesLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("MaxDurationMinutesLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("MinDurationMinutesLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("FailureCountLastWeek", NHibernateUtil.Int32)
            .AddScalar("SuccessCountLastWeek", NHibernateUtil.Int32)
            .AddScalar("FailurePercentageLastWeek", NHibernateUtil.Int32)
            .AddScalar("FailureCountLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("SuccessCountLast4Weeks", NHibernateUtil.Int32)
            .AddScalar("FailurePercentageLast4Weeks", NHibernateUtil.Int32)
            .SetParameter("endDate", endDate)
            .SetResultTransformer(Transformers.AliasToBean<JobMetric>())
            .SetTimeout(180);

            return await query.ListAsync<JobMetric>();
        }
    }
}
