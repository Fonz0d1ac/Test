-- Find the rows that caused:
--   ('42000', '... Error converting data type nvarchar to float. (8114)')
-- [Earned hour] is an nvarchar column, so it can hold text that is not a number.
-- One such row in the 30-day window aborted the whole StationModel query.

-- 1) FG_Database_All — the table the failing query reads.
SELECT  [Earned hour]              AS bad_value,
        COUNT(*)                   AS rows_affected,
        MIN([Production date])     AS first_seen,
        MAX([Production date])     AS last_seen
FROM    [dbo].[FG_Database_All]
WHERE   CAST([Production date] AS DATE) >= DATEADD(day, -30, CAST(GETDATE() AS DATE))
  AND   [Earned hour] IS NOT NULL
  AND   LTRIM(RTRIM(CAST([Earned hour] AS nvarchar(50)))) <> ''
  AND   TRY_CAST([Earned hour] AS float) IS NULL
GROUP BY [Earned hour]
ORDER BY rows_affected DESC;

-- 2) The individual rows, so you can correct them at source.
SELECT  [Production date], [Input time], Station, PO, [Part no],
        [License Plate], [Earned hour]
FROM    [dbo].[FG_Database_All]
WHERE   CAST([Production date] AS DATE) >= DATEADD(day, -30, CAST(GETDATE() AS DATE))
  AND   [Earned hour] IS NOT NULL
  AND   LTRIM(RTRIM(CAST([Earned hour] AS nvarchar(50)))) <> ''
  AND   TRY_CAST([Earned hour] AS float) IS NULL
ORDER BY [Production date] DESC, [Input time] DESC;

-- 3) Same check on the WIP table — it has the same nvarchar column and the same
--    exposure (two other queries in combined_dashboard.py read it).
SELECT  [Earned hour] AS bad_value, COUNT(*) AS rows_affected
FROM    [dbo].[Nhaplecuoingay_All]
WHERE   CAST([Production date] AS DATE) >= DATEADD(day, -30, CAST(GETDATE() AS DATE))
  AND   [Earned hour] IS NOT NULL
  AND   LTRIM(RTRIM(CAST([Earned hour] AS nvarchar(50)))) <> ''
  AND   TRY_CAST([Earned hour] AS float) IS NULL
GROUP BY [Earned hour]
ORDER BY rows_affected DESC;
