# SQL Server Snippets

## PostgreSQL difference between two dates
```sql

(DATE_PART('year',"end date" ) - DATE_PART('year', "start date"))||'  Years'||' '|| (date_part('month',"end date")-date_part('month',"start date"))||'  Months'  as "Tenure"

```

## Date in UK format
```sql
Select SQL Server date formats
;CONVERT (char(10),Getdate(),103)
```
## SQL SERVER VARIABLE
```sql
GO
DECLARE 
@iVariable INT = 1
,@vVariable VARCHAR(100) = 'myvar'
,@dDateTime DATETIME = GETDATE()
SELECT @iVariable iVar, @vVariable vVar, @dDateTime dDT
GO
```
## Camelcase UDF
```sql
USE [DM_1457_BusScience];
GO
SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO
ALTER FUNCTION [CTL].[udf_CleanField_CamelCase]
(
 @InputFieldRecord VARCHAR(8000)
 )
RETURNS VARCHAR(8000)
AS

BEGIN
DECLARE @OutputFieldRecord VARCHAR(8000)

 -- Trim Data
 SET @OutputFieldRecord = LTRIM(RTRIM(@InputFieldRecord))

 -- Double Spaces to single spaces
 IF @OutputFieldRecord LIKE '%  %' -- double spaces
 BEGIN
 SET @OutputFieldRecord = REPLACE(@OutputFieldRecord, '  ', ' ')
 END

 -- To Title Case
 DECLARE @Reset bit;
 DECLARE @ProcessFieldRecord varchar(8000);
 DECLARE @i int;
 DECLARE @c char(1);

 SELECT @Reset = 1, @i=1, @ProcessFieldRecord = '';

 WHILE (@i <= LEN(@OutputFieldRecord))
 SELECT @c= SUBSTRING(@OutputFieldRecord, @i, 1),
 @ProcessFieldRecord = @ProcessFieldRecord +
	CASE WHEN @Reset=1 THEN UPPER(@c) ELSE LOWER(@c) END,
 @Reset = CASE WHEN @c like '[a-zA-Z]' THEN 0 ELSE 1 END,
 @i = @i +1

 SET @OutputFieldRecord = @ProcessFieldRecord
 RETURN @OutputFieldRecord
END
GO
```

## Remove Duplicates
```sql
--region remove duplicates from DW_ODS.tbl_ff_comp_pillar_mappings
WITH MUNGE1
			AS (SELECT
						[Raw Data Name]
					 ,[Mapped Name]
					 ,RN =
							ROW_NUMBER ()
							OVER
							(
								PARTITION BY [Raw Data Name], [Mapped Name]
								ORDER BY [Raw Data Name]
							)
					FROM
						DW_ODS.TBL_FF_COMP_PILLAR_MAPPINGS)
--check -SELECT 	MUNGE1.*FROM	MUNGE1  where RN>1
--move to new table - select * into DW_ODS.tbl_ff_comp_pillar_mappings from munge1
--delete--
delete from munge1 where rn>1
--check - select *from DW_ODS.tbl_ff_comp_pillar_mappings
;
;
--endregion
```
