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
	  WHERE (dept.NAME LIKE '%Video%')
	  AND (shr.StdDate BETWEEN '<<start>>' AND '<<end>>') 
	  AND (shr.ShrinkCategory IN ('Scheduled', 'Unplanned OOO'))
	  AND ([ShrinkCode] <> 'STF-MGMT-OVR UNPAID')
	  GROUP BY shr.StdDate, shr.EmpID, shr.ShrinkCategory) as Shrink_Table
PIVOT(
	SUM([Shrink (sec)])
	FOR [ShrinkCategory] IN ([Unplanned OOO], [Scheduled])
	) AS piv
ORDER BY [StdDate] DESC, EmpID ASC;