USE [Kaplan_2015]
GO

/****** Object:  View [kaplan].[Kaplan_Scheduler_Tasklist]    Script Date: 02/07/2015 10:20:48 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


ALTER VIEW [kaplan].[Kaplan_Scheduler_Tasklist]
AS
-- the week commencing calendar date has to be derived from ISO Week number used in Celcat where the date is masked
WITH SevenDayCalendar AS
(
-- find the first Monday of the year
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 1) AS CalendarDate
UNION 
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 2)
UNION
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 3)
UNION 
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 4)
UNION 
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 5)
UNION 
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 6)
UNION 
SELECT DATEFROMPARTS(DATEPART(year, GETDATE()), 1, 7)
), 
WeekCommencing AS
(
-- determine which ISO week number this Monday falls into
SELECT CalendarDate
, DATEPART (ISO_WEEK,CalendarDate) AS isoweek
FROM SevenDayCalendar
WHERE DATEPART(DW,CalendarDate) = 2 -- Monday
-- recursively add 7 days up to 53 weeks
UNION ALL
SELECT DATEADD(week, 1, CalendarDate), isoweek + 1
FROM WeekCommencing
WHERE isoweek < 53
),
EventWeekNumber AS
(
SELECT e1.Event_id, 
CHARINDEX('Y',e1.weeks) AS event_week
FROM [dbo].[CT_EVENT] e1
WHERE CHARINDEX('Y',e1.weeks) > 0
UNION ALL
SELECT e2.Event_id, 
CHARINDEX('Y',e2.weeks, ewn.event_week + 1) AS event_week
FROM [dbo].[CT_EVENT] e2
INNER JOIN EventWeekNumber ewn
ON e2.event_id = ewn.event_id
AND CHARINDEX('Y',e2.weeks, ewn.event_week + 1) > 0
)
/* base data */
SELECT DISTINCT ewn.event_id
, e.[event_name]
, CASE 
		WHEN e.[day_of_week] = 0 THEN 'Monday'
		WHEN e.[day_of_week] = 1 THEN 'Tuesday'
		WHEN e.[day_of_week] = 2 THEN 'Wednesday'
		WHEN e.[day_of_week] = 3 THEN 'Thurday'
		WHEN e.[day_of_week] = 4 THEN 'Friday'
		WHEN e.[day_of_week] = 5 THEN 'Saturday'
		WHEN e.[day_of_week] = 6 THEN 'Sunday'
	END AS event_day
, ewn.event_week
, CONVERT(datetime, wc.CalendarDate) AS week_commencing -- excel does not recognise date type as date
, e.[start_time] AS event_start_time
, e.[end_time] AS event_end_time
, ec.[name] AS event_category
, COALESCE(e.[capacity_req],g.[group_size],0) AS pax
, d.name AS department_name
, m.[custom2] AS module_university
, e.[notes] event_notes
--, m.[unique_name] AS module_unique_name
, m.[name] AS module_name
--, g.unique_name AS MID_name
, m.[custom3] AS module_level
, r.unique_name AS room_name
, s.unique_name AS staff_name
FROM EventWeekNumber ewn
INNER JOIN WeekCommencing wc
ON ewn.event_week = wc.isoweek
INNER JOIN [dbo].[CT_EVENT] e
ON ewn.event_id = e.event_id
LEFT OUTER JOIN [dbo].[CT_EVENT_CAT] ec
ON e.[event_cat_id] = ec.event_cat_id
INNER JOIN [dbo].[CT_EVENT_MODULE] em
ON e.event_id = em.event_id
INNER JOIN [dbo].[CT_MODULE] m
ON em.module_id = m.module_id
LEFT OUTER JOIN [dbo].[CT_EVENT_ROOM] er
ON e.event_id = er.event_id
LEFT OUTER JOIN [dbo].[CT_ROOM] r
ON er.room_id = r.room_id
LEFT OUTER JOIN [dbo].[CT_DEPT] d
ON e.dept_id = d.dept_id
LEFT OUTER JOIN [dbo].[CT_EVENT_STAFF] es
ON e.event_id = es.event_id
LEFT OUTER JOIN [dbo].[CT_STAFF] s
ON es.staff_id = s.staff_id
LEFT OUTER JOIN [dbo].[CT_EVENT_GROUP] eg
ON e.event_id = eg.event_id
LEFT OUTER JOIN [dbo].[CT_GROUP] g
ON eg.group_id = g.group_id
WHERE e.suspended = 'N'

