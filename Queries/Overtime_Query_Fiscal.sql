declare @tday date = dateadd(hour,-5,getutcdate());
declare @currentFiscal date = 
CASE
	WHEN DATEPART(dd,@tday) < 29 THEN DATEADD(d,28-DATEPART(dd,@tday),@tday)
	ELSE DATEADD(m,1,DATEADD(d,-1*(DATEPART(dd,@tday)-28),@tday))
END;
;
declare @fMonth date = datefromparts(year(@currentFiscal),month(@currentFiscal),1);

declare @lastFM date = DATEADD(MONTH, -1, @fMonth)
declare @lookbackFM date = DATEADD(MONTH, -1, @lastFM)

DECLARE @sDate datetime;
DECLARE @eDate datetime;
;
SET @sDate = datefromparts(year(DATEADD(MONTH, -1, @lookbackFM)),month(DATEADD(MONTH, -1, @lookbackFM)),29);
SET @eDate = datefromparts(year(@lastFM),month(@lastFM),28);
;
DECLARE @sNom decimal(15,1);
DECLARE @eNom decimal(15,1);
;
SET @sNom = DATEDIFF(d,'12/30/1899',@sDate);
SET @eNom = DATEDIFF(d,'12/30/1899',@eDate);

SELECT FiscalMonth
	  ,PSID
	  ,SUM(DURATION) / 60 AS [OT Total]
FROM (
		SELECT	DISTINCT agt.EmployeeID AS PSID
				,dt.FiscalMonth
				,seg.SEG_CODE AS Segment
				,(sched.STOP_MOMENT - sched.START_MOMENT) as [Duration]
		FROM [Aspect].[WFM].[DET_SEG] as sched WITH(NOLOCK)
		INNER JOIN [DimensionalMapping].[DIM].[SEG_CODE] as seg WITH(NOLOCK)
		ON sched.SEG_CODE_SK = seg.SEG_CODE_SK
		INNER JOIN [DimensionalMapping].[DIM].[Agent] AS agt WITH(NOLOCK)
		ON sched.EMP_SK = agt.EMP_SK
		INNER JOIN [UXID].[EMP].[Workers] AS ros with(NOLOCK)
		ON REPLACE(agt.[EmployeeID],' ','') = REPLACE(ros.[NETIQWORKERID], ' ', '')
		INNER JOIN [UXID].[REF].[Departments] AS dept WITH(NOLOCK)
		ON ros.DEPARTMENTID = dept.DEPARTMENTID
		INNER JOIN [DimensionalMapping].[DIM].Date_Table AS dt WITH(NOLOCK)
		ON sched.NOM_DATE = dt.NomDate
		WHERE (dept.NAME IN ('Resi Video Repair Call Ctrs', 'Resi Video Repair CC'))
		AND (sched.NOM_DATE BETWEEN @sNom AND @eNom)
		AND (seg.SEG_CODE = 'FLEX_TIME_GIVE_BACK' OR seg.SEG_CODE = 'OVERTIME' OR seg.SEG_CODE = 'CANCEL_OVERTIME')
) AS subquery
GROUP BY FiscalMonth, PSID
ORDER BY FiscalMonth, PSID;