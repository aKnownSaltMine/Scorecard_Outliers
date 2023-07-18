SET NOCOUNT ON;
;
DECLARE @FiscalStart datetime;
DECLARE @FiscalNom numeric(15,1);
SET @FiscalStart = (CASE
              WHEN DATEPART(dd,GETDATE()) < 29 THEN DATEADD(d,28-DATEPART(dd,GETDATE()),GETDATE())
              ELSE DATEADD(m,1,DATEADD(d,-1*(DATEPART(dd,GETDATE())-28),GETDATE()))END);
;
/* 032621 - Updated FiscalStart to reflect the acutal fiscal Start */
SET @FiscalStart = DateAdd(d,1,DateAdd(m,-1,@FiscalStart));
SET @FiscalNom = DATEDIFF(d,'12/30/1899',DateAdd(m,-1,@FiscalStart));

;
WITH peer as (
	SELECT e.ID
		,p.CODE [Peer]
		,ROW_NUMBER() OVER(PARTITION by e.EMP_SK ORDER BY p.START_NOM_DATE DESC)peerOrd
	FROM GVPOperations.VID.ATT_EMP e with(nolock)
	INNER JOIN GVPOperations.VID.ATT_Peer p with(nolock) on e.EMP_SK=p.EMP_SK
	/*Stop showing peer once someone crosses out of the fiscal checked against UXID*/
	WHERE e.TERM_NOM_DATE >= @FiscalNom
),staff as (
	SELECT e.ID
		,p.CODE [Staff]
		,ROW_NUMBER() OVER(PARTITION by e.EMP_SK ORDER BY p.START_NOM_DATE DESC)staffOrd
	FROM GVPOperations.VID.ATT_EMP e with(nolock)
	INNER JOIN GVPOperations.VID.ATT_Staff p with(nolock) on e.EMP_SK=p.EMP_SK
	/*Stop showing peer once someone crosses out of the fiscal checked against UXID*/
	WHERE e.TERM_NOM_DATE >= @FiscalNom
)
SELECT w.SAMACCOUNTNAME,w.ENTITYACCOUNT PID
	,s.FIRSTNAME+ ' ' + s.LASTNAME BossName
	,m.FIRSTNAME+ ' ' + m.LASTNAME BossBossName
	,w.FIRSTNAME+ ' ' + w.LASTNAME EmpName
	,jc.TITLE EmpTitle
	,ISNULL(p.Peer,'')[Peer]
	,l.STATE+' '+l.CITY WorkLocation
	,w.NETIQWORKERID
	,case when w.STATUSID = 1 then 'Active' when (w.STATUSID = 2 or w.STATUSID = 3) then 'LOA' else '' end STATUSID
	,format(w.TERMINATEDDATE,'MM/dd/yyyy') [TERMINATEDDATE]
/*022421: Changed to Service Date to properly align with how Tenure and employee dates should be reported;*/
	,format(w.SERVICEDATE,'MM/dd/yyyy') [HIREDATE]
	,ISNULL(st.Staff,'')[Staff]
	/*010520: Removed to avoid dependency on a traffic table*/
	/*,case when currdate.CURRENTPOSITIONDATE IS NULL OR currdate.CURRENTPOSITIONDATE = 'NULL' THEN NULL ELSE currdate.CURRENTPOSITIONDATE END CURRENTPOSITIONDATE*/
	,w.CURRENTPOSITIONSTARTDATE CURRENTPOSITIONDATE
	/*040121:	Added to deal with Don Haskins being added*/
	,REPLACE(ma.MGMTAREANAME,'CC-','') MGMTAREANAME
	,sch.[Days Worked]
	,sch.[Start/Stop]
	,his.StartDate as [WP Start Date]
	,wp.WorkPlace
FROM UXID.EMP.Workers w with(nolock)
INNER JOIN UXID.REF.Departments d with(nolock) on w.DEPARTMENTID=d.DEPARTMENTID
LEFT OUTER JOIN UXID.EMP.Workers s with(nolock) on w.SUPERVISORID=s.WORKERID
LEFT OUTER JOIN UXID.EMP.Workers m with(nolock) on s.SUPERVISORID=m.WORKERID
LEFT OUTER JOIN UXID.REF.Job_Codes jc with(nolock) on w.JOBCODEID=jc.JOBCODEID
LEFT OUTER JOIN UXID.REF.Locations l with(nolock) on w.LOCATIONID=l.LOCATIONID
LEFT OUTER JOIN UXID.REF.Management_Areas ma with(nolock) on w.MANAGEMENTAREAID=ma.MANAGEMENTAREAID
LEFT OUTER JOIN peer p on p.peerOrd=1 AND p.ID=w.NETIQWORKERID
LEFT OUTER JOIN staff st on st.staffOrd=1 AND st.ID=w.NETIQWORKERID
LEFT OUTER JOIN GVPOperations.VID.PROD_SCHED sch with(nolock) on w.NETIQWORKERID = sch.PSID
LEFT OUTER JOIN [DimensionalMapping].DIM.Agent AS agt WITH(NOLOCK) ON CAST(REPLACE(w.NETIQWORKERID,' ', '') AS varchar)=REPLACE(agt.EmployeeID, ' ','')
LEFT OUTER JOIN [DimensionalMapping].[DIM].[Historical_Agent_Work_Place] AS his WITH(NOLOCK) ON	his.EMP_SK=agt.EMP_SK
LEFT OUTER JOIN [DimensionalMapping].[DIM].[Work_Place] AS wp WITH(NOLOCK) ON his.Work_Place_SK=wp.Work_Place_SK
/*010520: Removed to avoid dependency on a traffic table*/
/*LEFT OUTER JOIN UXID.REF.PROD_CURRENTPOSDATE_temp currdate on w.NETIQWORKERID = currdate.PEOPLESOFTID*/
/*
	102220 - Per Jincky, Wants to include 
	022221 - Updated to include Terms within One Year
	030121 - Reverted TermDate to 45 days per feedback from Michele
*/
WHERE (w.TERMINATEDDATE IS NULL OR w.TERMINATEDDATE >= DateAdd(d,-45,@FiscalStart)) AND 
	((CASE
		WHEN d.NAME LIKE '%resi%Video%' THEN 1
		/*091922: Tweaked from Cust Serv to capture abbreviated Dept*/
		WHEN d.NAME LIKE 'Cust%S%v%Train%' THEN 1
		/*010521: Added per Brandon*/ 
		WHEN d.NAME LIKE 'Cust%HR' THEN 1
		WHEN d.NAME LIKE 'Cust%Recruitmt' THEN 1
		ELSE 0
	END) = 1 or
/* 051822 - As Per Jincky, Added McAllen to EHH Roster */
	(l.STATE+' '+l.CITY like 'TX McAllen')
	)
	AND (CASE
			WHEN his.EndDate = '9999-12-31' THEN 1
			WHEN his.EndDate IS NULL THEN 1
			ELSE 0
		END)=1
order by 1