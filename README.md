# Balanced Scorecard Estimator

## Project Overview

In this particular project we are using Excel in cobination with SQL sources to pull multiple indicies and generate some index scores based on Actual vs. Planned percentages.  The complexity of this project is mostly in the overall scope and size of the data pulls, the granularity of information 
and a tiered scoring system that only awards points if certain scoring thresholds are achieved.  We also need to have data available in Week Ending / Monthly formats for year long availability.  Some VBA was utilized as well.  This Scorecard, once set up was autonomous and updates, postes to an available FTP site
daily without intervention after completion.  It was built in a manner that yearly updates to elements and tiered scoring were easily implemented.

## Data Sources

On Prem T-SQL database, moderately normalized

## Tools
- Excel   | Data Presentation to End User
- T-Sql   | Data Acqusition

## Data Acquisition / Preperation

- In order to maintain visible granualrity for the end users, a rollup approach was used.  All of the finer points of data were accumulated and each of those points would rollup to created the larger segmented groups.
- Three injection tables were created on the database
  -    Matrix_Raw
  -    Matrix_Raw_Division
  -    Matrix_Raw_District

---
- The base Matrix_Raw table would contain our lowest level of data and allow for further analysis and use in the future if any results were required outside of the initial scope.  This table would contain our complete hierarchal breakdown.  there is nothing outstanding about this base table other than it's going to be the base layer of all applicable data moving forward.  This base table is the building block for all analysis moving forward.
~~~~
CREATE TABLE [rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW]
(
	[REC_ID]			[int] IDENTITY(1,1) NOT NULL,
	[TY_DAY_DT]			[date] NULL,
	[TY_WND_DT]			[date] NULL,
	[PKG_Week]			[char](2) NULL,
	[MO_NUM]			[char](2) NOT NULL,
	[MO_NAME]			[nvarchar](30) NULL,
	[QTR_NR]			[char](2) NOT NULL,
	[TY_Year]			[char](4) NOT NULL,
	[REG_NR]			[varchar](4) NULL,
	[REG_NA]			[varchar](30) NULL,
	[REGION]			[varchar](37) NULL,
	[OP_GRP_NR]			[varchar](4) NULL,
	[OP_GRP_NA]			[varchar](25) NULL,
	[DIS_NR]			[varchar](4) NULL,
	[DIS_NA]			[varchar](30) NULL,
	[DISTRICT]			[varchar](37) NULL,
	[DIV_NR]			[varchar](4) NULL,
	[DIV_NA]			[varchar](12) NULL,
	[DIVISION]			[varchar](19) NULL,
	[CTR_NR]			[varchar](6) NULL,
	[CTR_NA]			[varchar](35) NULL,
	[BLD_NR]			[varchar](5) NULL,
	[BLD_NA]			[varchar](40) NULL,
	[BUILDING]			[varchar](49) NULL,
	[ID_1_Element_ID]		[int] NOT NULL,
	[ID_1_Volume]			[int] NOT NULL,
	[ID_1_Errors]			[int] NOT NULL,
	[ID_2A_Element_ID]		[int] NOT NULL,
	[ID_2A_Volume]			[int] NOT NULL,
	[ID_2A_Errors]			[int] NOT NULL,
	[ID_2B_Element_ID]		[int] NOT NULL,
	[ID_2B_Volume]			[int] NOT NULL,
	[ID_2B_Errors]			[int] NOT NULL,
	[ID_2C_Element_ID]		[int] NOT NULL,
	[ID_2C_Volume]			[int] NOT NULL,
	[ID_2C_Errors]			[int] NOT NULL,
	[ID_2D_Element_ID]		[int] NOT NULL,
	[ID_2D_Volume]			[int] NOT NULL,
	[ID_2D_Errors]			[int] NOT NULL,
~~~~

  This scope was carried out across many more trackables...

---

  In our rollup tables for Division and District we encorporated some other Matrix elements for our reporting calculations and left out some of the hierarchal granularity becasue we will have summarized data contained within

~~~~
CREATE TABLE [rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIV]  
(  
	[TY_WND_DT]              [date] NULL,  
	[PKG_WEEK]               [char](2) NULL,  
	[MO_NUM]                 [char](2) NULL,  
	[REGION]                 [varchar](37) NULL,  
	[District]               [varchar](37) NULL,  
	[Division]               [varchar](19) NULL,  
	[OP_Grp_NR]              [nchar](10) NULL,  
	[OP_Grp_NA]              [nchar](50) NULL,  
	[ID_1_Element_ID]        [int] NULL,  
	[ID_1_Volume]            [int] NULL,  
	[ID_1_Errors]            [int] NULL,  
	[ID_1_Freq]              [int] NULL,  
	[ID_1_Goal]              [int] NULL,  
	[ID_1_Eff]               [float] NULL,  
	[ID_1_Points]            [float] NULL,  
	[ID_1_Possible_Points]   [int] NULL,  
~~~~
 Notice that we are including the Volume and Errors fields (Numerator- Denominator for our percentages to generate our Frequency) as well as a Goal, Effective, Points and Possible Points fields 
 
 Again - This scope was carried out across many more trackables...

 ---

Once our baseline tables were ready to go, it was time to start some date manipulations.  Due to the amount of data and the fact that we need to retain information for the year in our final reporting product we don't want to max out server resources as the year moves forward regenerating previous information across all levels every time the data was pulled - for example when December rolls around we don't want to regenerate all information from January to December - we will append information to our baseline tables instead.  

The server data did contain some baseline calendar tables but we manipulated a few things in order to suit our needs.  

We set up a stored procedue that can be fired via Powershell later on    

We do use a function created on the server as well to find dates before and after our selected date referred to as [fn_Calendar_TY_LY]

~~~~
CREATE PROCEDURE [rpt].[sp_BB_MATRIX_SETUP_CURRENT_WEEKENDING]  

AS

BEGIN

-- Begin Variable Setup
--------------------------------------------------------------------------------------------------------------------
Declare @Current_WE				as Date
Declare @Staged_Date				as Date
Declare @Matrix_Max_Date			as Date
Declare @CMO					as nvarchar(2)
Declare @Current_Year				as nvarchar(4)
Declare @DIS_Matrix_Max_MTH			as nvarchar(2)
Declare @DIV_Matrix_Max_MTH			as nvarchar(2)


-- Find Maximun date set in raw matrix file
Set @Matrix_Max_Date	= (SELECT MAX([TY_WND_DT]) FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW])

-- Find latest Month in District Raw File for Monthly Setups / Adds
Set @DIS_Matrix_Max_MTH = (SELECT MAX([MO_NUM]) FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW])

-- Finds out if there are Monthly slots created in the District results table so we don't overwrite if so
Set @DIV_Matrix_Max_MTH = (SELECT MAX([MO_NUM]) FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW])

-- Find current week ending
Set @Current_WE	=	(SELECT Top 1 WeekEndDate_TY
			FROM [DADH1001].[src].[fn_Calendar_TY_LY] (GETDATE(),0,0) --START DATE, WEEKS BEFORE, WEEKS AFTER  
			WHERE [DayDate_TY]<GETDATE()-1
			)

-- Find current Month
Set @CMO	=	(Select Top 1 monthnumber from  [DADH1001].[src].[t_corpcodes_calendar]
			Where WeekEndDate = @Current_WE
			)
									
-- Find current Year
Set @Current_Year=	(Select Top 1 YearNumber from  [DADH1001].[src].[t_corpcodes_calendar]
			Where WeekEndDate = @Current_WE
			)

-- End Variable Setup
------------------------------------------------------------------------------------------------------

~~~~

---

Once we established what our current week status was - we used some temp tables to set up TY/LY Calendar table.  
This gave us a complete reference table for all date possibilities across our Matrix table.

~~~~
-- Begin Calendar Dates

IF @Matrix_Max_Date < @Current_WE 
BEGIN

If Object_ID('tempdb..#CAL_TY') is not null 
BEGIN
Drop Table #CAL_TY
END
--------------------------------------------------
Select * into #CAL_TY
FROM
(
Select
						 
	  convert(datetime, DayDate,23) as TY_DAY_DT
	, DOW_CD as TY_DOW
	, DOW_NA as TY_DOW_NA
	, convert(datetime, WeekEndDate,23) as TY_WND_DT
	, pkgweeknumber as PKG_Week
	, monthnumber as MO_NUM
	, DATENAME(MONTH,Dateadd(MONTH,cast(MonthNumber as int),'2020-12-01'))as MO_NAME
	, quarternumber as QTR_NR
	, yearnumber as TY_Year
	, OperatingDayInd as OP_Day

FROM	 [DADH1001].[src].[t_corpcodes_calendar]
WHERE	 yearnumber=year (getdate()) 
						
) CAL_TY

If Object_ID('tempdb..#CAL_LY') is not null 
BEGIN
Drop Table #CAL_LY
END
--------------------------------------------------
	Select * into #CAL_LY
	FROM
	(
		Select   
				convert(datetime, DayDate,23) as LY_DAY_DT
			, DOW_CD as LY_DOW
			, DOW_NA as LY_DOW_NA
			, convert(datetime, WeekEndDate,23) as LY_WND_DT
			, pkgweeknumber as PKG_Week
			, monthnumber as MO_NUM
			, DATENAME(MONTH,Dateadd(MONTH,cast(MonthNumber as int),'2020-12-01'))as MO_NAME
			, quarternumber as QTR_NR
			, yearnumber as Year
			, OperatingDayInd as OP_Day

	FROM [DADH1001].[src].[t_corpcodes_calendar]
	where yearnumber= year (Dateadd(year, -1,getdate()))
	) CAL_LY
--------------------------------------------------
			
If Object_ID('tempdb..#CALLY') is not null 
BEGIN
Drop Table #CALLY
END

SELECT * INTO #CALLY
FROM (
			Select 
                         TY_WND_DT		
			,TY_DAY_DT	
			,TY_DOW	
			,TY_DOW_NA
			,LY_WND_DT		
			,LY_DAY_DT	
			,#CAL_TY.OP_Day
			,#CAL_TY.PKG_Week	
			,#CAL_TY.MO_Num	
			,#CAL_TY.MO_Name
			,#CAL_TY.QTR_NR
			,#CAL_TY.TY_Year
					
		FROM #CAL_TY
		Inner Join	
			#CAL_LY on #CAL_TY.PKG_Week = #CAL_LY.pkg_week
					AND TY_DOW = LY_DOW
					AND TY_WND_DT = @Current_WE		
					
		) ALL_CAL

-- End Calendar Dates
------------------------------------------------------------------------------------------------------
~~~~
----
This resulted in a complete breakout of our requred dates for the year
~~~~

TY_WND_DT	TY_DAY_DT	TY_DOW	TY_DOW_NA	LY_WND_DT	LY_DAY_DT	OP_Day	PKG_Week	MO_Num	MO_Name	QTR_NR	TY_Year
1/6/2024	12/31/2023	1	Sunday		1/7/2023	1/1/2022	1	1	1	January	1	2024
1/6/2024	1/1/2024	2	Monday		1/7/2023	1/2/2022	2	1	1	January	1	2024
1/6/2024	1/2/2024	3	Tuesday		1/7/2023	1/3/2022	3	1	1	January	1	2024
1/6/2024	1/3/2024	4	Wednesday	1/7/2023	1/4/2022	4	1	1	January	1	2024
1/6/2024	1/4/2024	5	Thursday	1/7/2023	1/5/2022	5	1	1	January	1	2024
1/6/2024	1/5/2024	6	Friday		1/7/2023	1/6/2022	6	1	1	January	1	2024
1/6/2024	1/6/2024	7	Saturday	1/7/2023	1/7/2022	7	1	1	January	1	2024
1/13/2024	1/7/2024	1	Sunday		1/14/2023	1/8/2022	1	2	1	January	1	2024
1/13/2024	1/8/2024	2	Monday		1/14/2023	1/9/2022	2	2	1	January	1	2024
1/13/2024	1/9/2024	3	Tuesday		1/14/2023	1/10/2022	3	2	1	January	1	2024
1/13/2024	1/10/2024	4	Wednesday	1/14/2023	1/11/2022	4	2	1	January	1	2024
1/13/2024	1/11/2024	5	Thursday	1/14/2023	1/12/2022	5	2	1	January	1	2024
1/13/2024	1/12/2024	6	Friday		1/14/2023	1/13/2022	6	2	1	January	1	2024
1/13/2024	1/13/2024	7	Saturday	1/14/2023	1/14/2022	7	2	1	January	1	2024

etc...
~~~~




 
  

  
  	


