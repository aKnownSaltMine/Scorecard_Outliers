/****** Script for SelectTopNRows command from SSMS  ******/
declare @tday date = dateadd(hour,-5,getutcdate());
-- set @tday = datefromparts(2023,3,1)
declare @currentFiscal date = 
CASE
	WHEN DATEPART(dd,@tday) < 29 THEN DATEADD(d,28-DATEPART(dd,@tday),@tday)
	ELSE DATEADD(m,1,DATEADD(d,-1*(DATEPART(dd,@tday)-28),@tday))
END;
;

declare @fMonth date = datefromparts(year(@currentFiscal),month(@currentFiscal),1);
declare @lookbackFM date = DATEADD(MONTH, -2, @fMonth)

DECLARE @sDate datetime = 
CASE
	WHEN DATEPART(mm, @lookbackFM) = 3 AND ((DATEPART(yy, @lookbackFM) % 4) <> 0 OR (((DATEPART(yy, @lookbackFM) % 100) = 0) OR ((DATEPART(yy, @lookbackFM) % 400) = 0)))
	THEN @fmonth
	ELSE datefromparts(year(DATEADD(MONTH, -1, @lookbackFM)),month(DATEADD(MONTH, -1, @lookbackFM)),29)
END;
DECLARE @eDate datetime = datefromparts(year(@fMonth),month(@fMonth),28);
;

set @sDate = DATEADD(day, -120, @fMonth);


Select StdDate AS [Date]
	  ,EmpID
	  ,[Unplanned OOO]
	  ,[Scheduled]
FROM (SELECT shr.StdDate
		  ,[EmpID]
		  ,[ShrinkCategory]
		  ,Sum([ShrinkSeconds]) as [Shrink (sec)]
	  FROM [Aspect].[WFM].[BI_Daily_CS_Shrinkage] as shr
	  INNER JOIN [UXID].[EMP].[Workers] AS ros with(NOLOCK)
	  ON REPLACE(shr.[EmpID],' ','') = REPLACE(ros.[NETIQWORKERID], ' ', '')
	  INNER JOIN [UXID].[REF].[Departments] AS dept WITH(NOLOCK)
	  ON ros.DEPARTMENTID = dept.DEPARTMENTID
	  WHERE (dept.NAME IN ('Resi Video Repair Call Ctrs', 'Resi Video Repair CC'))
	  AND (shr.StdDate BETWEEN @sDate AND @eDate) 
	  AND (shr.ShrinkCategory IN ('Scheduled', 'Unplanned OOO'))
	  AND ([ShrinkCode] <> 'STF-MGMT-OVR UNPAID')
	  GROUP BY shr.StdDate, shr.EmpID, shr.ShrinkCategory) as Shrink_Table
PIVOT(
	SUM([Shrink (sec)])
	FOR [ShrinkCategory] IN ([Unplanned OOO], [Scheduled])
	) AS piv
ORDER BY [StdDate] DESC, EmpID ASC;