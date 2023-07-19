Select StdDate AS [Date]
	  ,EmpID
	  ,[Out of Center - Planned]
	  ,[Out of Center - Unplanned]
	  ,[Scheduled Hours]
FROM (SELECT shr.StdDate
		  ,[EmpID]
		  ,[ShrinkType]
		  ,Sum([ShrinkSeconds]) as [Shrink (sec)]
	  FROM [Aspect].[WFM].[BI_Daily_CS_Shrinkage] as shr
	  INNER JOIN [UXID].[EMP].[Workers] AS ros with(NOLOCK)
	  ON REPLACE(shr.[EmpID],' ','') = REPLACE(ros.[NETIQWORKERID], ' ', '')
	  INNER JOIN [UXID].[REF].[Departments] AS dept WITH(NOLOCK)
	  ON ros.DEPARTMENTID = dept.DEPARTMENTID
	  WHERE (dept.NAME LIKE '%Video%')
	  AND (shr.StdDate BETWEEN '<<start>>' AND '<<end>>') 
	  AND ((shr.ShrinkType LIKE '%Out of Center%') 
	  OR (shr.ShrinkType Like '%Scheduled%'))
	  GROUP BY shr.StdDate, shr.EmpID, shr.ShrinkType) as Shrink_Table
PIVOT(
	SUM([Shrink (sec)])
	FOR [ShrinkType] IN ([Out of Center - Planned], [Out of Center - Unplanned], [Scheduled Hours])
	) AS piv
ORDER BY [StdDate] DESC, EmpID ASC;