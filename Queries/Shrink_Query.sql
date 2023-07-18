/****** Script for SelectTopNRows command from SSMS  ******/
declare @tday date = dateadd(hour,-5,getutcdate());
declare @currentFiscal date = 
CASE
	WHEN DATEPART(dd,@tday) < 29 THEN DATEADD(d,28-DATEPART(dd,@tday),@tday)
	ELSE DATEADD(m,1,DATEADD(d,-1*(DATEPART(dd,@tday)-28),@tday))
END;
;
declare @fMonth date = datefromparts(year(@currentFiscal),month(@currentFiscal),1);

declare @lastFM date = DATEADD(MONTH, -1, @fMonth)
declare @lookbackFM date = DATEADD(MONTH, -3, @lastFM)

DECLARE @sDate datetime;
DECLARE @eDate datetime;
;
SET @sDate = datefromparts(year(DATEADD(MONTH, -1, @lookbackFM)),month(DATEADD(MONTH, -1, @lookbackFM)),29);
SET @eDate = datefromparts(year(@lastFM),month(@lastFM),28);

Select FiscalMonth
	  ,EmpID
	  ,ISNULL([Unplanned OOO],0) AS [Unplanned OOO]
	  ,[Scheduled]
FROM (SELECT fm.FiscalMonth
		  ,[EmpID]
		  ,[ShrinkCategory]
		  ,Sum([ShrinkSeconds]) as [Shrink (sec)]
	  FROM [Aspect].[WFM].[BI_Daily_CS_Shrinkage] as shr
	  INNER JOIN [UXID].[EMP].[Workers] AS ros with(NOLOCK)
	  ON REPLACE(shr.[EmpID],' ','') = REPLACE(ros.[NETIQWORKERID], ' ', '')
	  INNER JOIN [UXID].[REF].[Departments] AS dept WITH(NOLOCK)
	  ON ros.DEPARTMENTID = dept.DEPARTMENTID
	  INNER JOIN DimensionalMapping.DIM.Date_Table AS fm WITH(NOLOCK)
	  ON shr.StdDate=fm.StdDate
	  WHERE (dept.NAME IN ('Resi Video Repair Call Ctrs', 'Resi Video Repair CC'))
	  AND (shr.StdDate BETWEEN @sDate AND @eDate) 
	  AND (shr.ShrinkCategory IN ('Scheduled', 'Unplanned OOO'))
	  GROUP BY fm.FiscalMonth, shr.EmpID, shr.ShrinkCategory) as Shrink_Table
PIVOT(
	SUM([Shrink (sec)])
	FOR [ShrinkCategory] IN ([Unplanned OOO], [Scheduled])
	) AS piv
ORDER BY [FiscalMonth], EmpID;