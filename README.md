# Balanced Scorecard Estimator
----
## Project Overview

In this particular project we are using Excel in cobination with SQL sources to pull multiple indicies and generate some index scores based on Actual vs. Planned percentages.  The complexity of this project is mostly in the overall scope and size of the data pulls, the granularity of information 
and a tiered scoring system that only awards points if certain scoring thresholds are achieved.  We also need to have data available in Week Ending / Monthly formats for year long availability.  Some VBA was utilized as well.  This Scorecard, once set up was autonomous and updates, postes to an available FTP site
daily without intervention after completion.  It was built in a manner that yearly updates to elements and tiered scoring were easily implemented.
----

## Table of Contents
- [Tools](#Tools)
- [Data Acquisition / SQL Preperation](#Data-Acquisition)
- [Excel Configuration](#Excel-Configuration)


## Data Sources

On Prem T-SQL database, moderately normalized

## Tools
- Excel   | Data Presentation to End User
- T-Sql   | Data Acqusition

## Data Acquisition

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
----
Now that we have a calendar of all avalable dates possible - we set up our mega hierarchy table that contains our granular data information across each segment, for each date -
On occasion if changes occur to our granular levels - for example an added location - this mega table accounts for it  

We are using some server tables that contain proprietary information to set this up a bit in this code example

~~~~

-- Begin Hierarchy Setup
------------------------------------------------------------------------------------------------------
-- PKG HIERARCHY SETUP 
------------------------------------------------------------------------------------------------------

If Object_ID('tempdb..#PKG_HIERARCHY') is not null 
BEGIN
Drop Table #PKG_HIERARCHY
END

Select * INTO #PKG_HIERARCHY FROM 
(SELECT 
PKGCTR.REG_NR
,PKGCTR.REG_NA
,PKGCTR.OP_GRP_NR
,PKGCTR.OP_GRP_NA
,PKGCTR.DIS_NR
,PKGCTR.DIS_NA
,PKGCTR.DIV_NR
,PKGCTR.DIV_NA
,PKGCTR.BLD_NR
,PKGCTR.NS_BLD_NR
,PKGCTR.NS_BLD_NA
,PKGCTR.NS_BLD_MNE_NA
,PKGCTR.SLIC1
,PKGCTR.SLIC2
,PKGCTR.PKG_NAME
,PKGCTR.LEVEL
,PKGCTR.ACT_IR
,PKGCTR.AUTO_PRELOAD

FROM DADH1001.src.vw_CTE_PkgCtrHierarchy PKGCTR

WHERE LEVEL IN ('CENTER1', 'CENTER2') AND ACT_IR = 1
) PKG_HIERARCHYTEMP

If Object_ID('tempdb..#CALENDAR_HIERARCHY') is not null 
BEGIN
Drop Table #CALENDAR_HIERARCHY
END

Select * into #CALENDAR_HIERARCHY 
FROM 
(SELECT 
  TY_DAY_DT
, TY_WND_DT
, PKG_Week
, MO_NUM
, MO_NAME
, QTR_NR
, TY_Year
, REG_NR
, REG_NA
, REGION		= TRIM(REG_NR) + ' - ' + TRIM(REG_NA)
, OP_GRP_NR
, OP_GRP_NA
, DIS_NR
, DIS_NA
, DISTRICT		= TRIM(DIS_NR) + ' - ' + TRIM(DIS_NA)
, DIV_NR
, DIV_NA
, DIVISION		= TRIM(DIV_NR) + ' - ' + TRIM(DIV_NA)
, CTR_NR		= TRIM(SLIC1)
, CTR_NA		= TRIM(PKG_NAME)
, BLD_NR		= TRIM(NS_BLD_NR)
, BLD_NA		= TRIM(NS_BLD_NA)
, BUILDING		=	CASE [LEVEL]
					WHEN 'CENTER1' THEN TRIM(SLIC1)  + ' - ' + TRIM(PKG_NAME)
					WHEN 'CENTER2' THEN TRIM(BLD_NR) + ' - ' + TRIM(NS_BLD_NA)
					END
						
-- Begin Element Setups

-- Customer First Index	------------------------------------------------------------------------------	
-- Late								
		, ID_1_Element_ID	= 1		
		, ID_1_Volume		= 0
		, ID_1_Errors	= 0					
-- Visibility
		-- No Scan	
		, ID_2A_Element_ID	= 021	
		, ID_2A_Volume		= 0
		, ID_2A_Errors	= 0
		-- Delivery Scan	
		, ID_2B_Element_ID	= 022	
		, ID_2B_Volume		= 0
		, ID_2B_Errors	= 0	
		-- Bulk Delivery Scan
		, ID_2C_Element_ID	= 023	
		, ID_2C_Volume		= 0
		, ID_2C_Errors	= 0	
		-- SPSF Scan - Tracking Only	
		, ID_2D_Element_ID	= 024	
		, ID_2D_Volume		= 0
		, ID_2D_Errors	= 0

~~~~


 In this code example we are also priming our Matrix tables with some element ID's for reference to act as PK's when pushing information to them  

 We do similar table setups to our Division and District Matrix tables as well...

 ----

 In a new table we set up a matrix of information pertaining to all elements - the Element Number (PK), name, Goal, Scoring Tier and Ponits assigned)  
 This tale allows us to future proof changes so that we can modify things just here, and it will carry across.  

 We pull this element setup information into a firable sproc that procesees our raw data source, configures it across our hierarchy and updates  
 our granular Matrix table and we then perform frequency calculations, establish our goals and points in a temp table so that we can use it all  
 further down the line in the sproc code and calculate some completed data that is finalized and ready for Excel importing

 ~~~~
CREATE PROCEDURE [rpt].[sp_DD_ELEMENT_BSC_A_CUST_FIRST_ID_1_RESP_LATE]  

AS

BEGIN

Declare @Element_ID int
Declare @Element_Name nvarchar(200)
Declare @Element_Goal int
Declare @Element_Tier as int
Declare @Element_Points as int
Declare @EFF as float
Declare @Func_STRING nvarchar(50)

Set @Element_ID		= 1	
Set @Element_Name	= (Select Element_Name		FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_ELEMENTS] Where ELement_ID = @Element_ID)
Set @Element_Goal	= (Select Goal			FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_ELEMENTS] Where ELement_ID = @Element_ID)
Set @Element_Tier	= (Select Tier			FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_ELEMENTS] Where ELement_ID = @Element_ID)
Set @Element_Points	= (Select Element_Points	FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_ELEMENTS] Where ELement_ID = @Element_ID)

Set @Func_STRING	= '[DADH1001].[src].[fn_Tier_'+ cast(@Element_Tier as nvarchar)+']'

-------------------------------------------------------------------------------------------
-- RAW ELEMENT PULL FOR PUSH TO RAW MATRIX
-------------------------------------------------------------------------------------------
If Object_ID('tempdb..#TEMPER') is not null 
BEGIN
	Drop Table #TEMPER
END
Select * into #TEMPER
FROM 
(
Select   	   
          DayDate
	, FAC_LOC_NR
	, OGZ_NR
	, VOLUME				= PKRL_VOL
  	, ERRORS				= PKRL_SF
		  
	from DADH1002.OE_SRC.PKRL ELLY
	RIGHT Join [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW] MATRIX
	ON  ELLY.DayDate	= MATRIX.TY_DAY_DT
	AND ELLY.OGZ_NR		= MATRIX.CTR_NR  
	 	   
) WORK_IT
-------------------------------------------------------------------------------------------
-- PUSH RAW TO MATRIX BASED ON DATE AND CENTER -- RAW VOLUME AND ERRORS
-------------------------------------------------------------------------------------------
UPDATE [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW] 
 	
SET       ID_1_VOLUME		= VOLUME
	, ID_1_ERRORS		= ERRORS
	FROM #TEMPER
	WHERE
	TY_DAY_DT=DAYDATE AND CTR_NR=OGZ_NR
-------------------------------------------------------------------------------------------
-- BEGIN CALCULATIONS FOR SUMMARY TABLES
-------------------------------------------------------------------------------------------
If Object_ID('tempdb..#CALCULATE_A') is not null 
BEGIN
	Drop Table #CALCULATE_A
END

Select * INTO #CALCULATE_A
FROM 
	(
		SELECT 

		  Element_ID=@Element_ID
		, TY_WND_DT
		, REGION
		, DISTRICT
		, SUM(ID_1_VOLUME) as VOL
		, SUM(ID_1_ERRORS) as ERR
		, FREQ = CAST(case when SUM(ID_1_ERRORS)=0 then SUM(ID_1_VOLUME) 
			 Else SUM(ID_1_VOLUME)/SUM(ID_1_ERRORS) end as FLOAT) -- This Cast gives us Decimals in the Effective
		, GOAL = @Element_Goal
		, TIER = @Element_Tier
		, POSSIBLE_POINTS = @Element_Points
		FROM
		[DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW] 
		GROUP BY TY_WND_DT, REGION, DISTRICT
	) CALC_A

If Object_ID('tempdb..#CALCULATE_DIS') is not null 
BEGIN
	Drop Table #CALCULATE_DIS
END

	Select * INTO #CALCULATE_DIS
	FROM 
	(
		SELECT
		  Element_ID
		, TY_WND_DT
		, REGION
		, DISTRICT
		, VOL
		, ERR
		, FREQ
		, GOAL
		, EFF = Case When (FREQ/GOAL)>1 Then 1 Else (FREQ/GOAL) END
		, TIER
		, POSSIBLE_POINTS
		, ACTUAL_POINTS = CASE 
							WHEN @ELEMENT_TIER = 98 
							THEN ( [DADH1001].[src].[fn_Tier_98] ((FREQ/GOAL),@Element_Points))
							  
							WHEN @ELEMENT_TIER = 90 
							THEN ( [DADH1001].[src].[fn_Tier_90] ((FREQ/GOAL),@Element_Points))

							WHEN @ELEMENT_TIER = 80 
							THEN ( [DADH1001].[src].[fn_Tier_80] ((FREQ/GOAL),@Element_Points))

							else 0 end
							
			
		FROM #CALCULATE_A ) CALC_DIS


-------------------------------------------------------------------------------------------
-- PUSH Final Data to District Results List
-------------------------------------------------------------------------------------------

UPDATE DISTRICT_RESULTS
	 	
SET   
		DISTRICT_RESULTS.ID_1_Element_ID				= DIS.Element_ID
		, DISTRICT_RESULTS.ID_1_VOLUME					= DIS.VOL
		, DISTRICT_RESULTS.ID_1_ERRORS					= DIS.ERR
		, DISTRICT_RESULTS.ID_1_Freq					= DIS.Freq
		, DISTRICT_RESULTS.ID_1_GOAL					= DIS.Goal
		, DISTRICT_RESULTS.ID_1_Eff					= DIS.Eff
		, DISTRICT_RESULTS.ID_1_Points					= DIS.Actual_Points
		, DISTRICT_RESULTS.ID_1_Possible_Points				= DIS.Possible_Points
		 
	FROM #CALCULATE_DIS DIS
	INNER JOIN [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIS] DISTRICT_RESULTS
	ON
	DIS.TY_WND_DT		= DISTRICT_RESULTS.TY_WND_DT 
	AND 
	DIS.REGION		= DISTRICT_RESULTS.REGION 
	AND
	DIS.DISTRICT		= DISTRICT_RESULTS.DISTRICT

------------------------------------------------------------------------------------------------------
-- DIVISION
------------------------------------------------------------------------------------------------------

If Object_ID('tempdb..#CALCULATE_B') is not null 
	BEGIN
		Drop Table #CALCULATE_B
	END

	Select * INTO #CALCULATE_B
	FROM 
		
		(
			SELECT 

			  ELement_ID=@Element_ID
			, TY_WND_DT
			, REGION
			, DISTRICT
			, DIVISION
			, SUM(ID_1_VOLUME) as VOL
			, SUM(ID_1_ERRORS) as ERR
			, FREQ = CAST(case when SUM(ID_1_ERRORS)=0 then SUM(ID_1_VOLUME) 
				else SUM(ID_1_VOLUME)/SUM(ID_1_ERRORS) end as FLOAT) -- This Cast gives us Decimals in the Effective
			, GOAL = @Element_Goal
			, TIER = @Element_Tier
			, POSSIBLE_POINTS = @Element_Points
			
			FROM
			[DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RAW] 
		
			GROUP BY TY_WND_DT, REGION, DISTRICT, DIVISION

		) CALC_B



	If Object_ID('tempdb..#CALCULATE_DIV') is not null 
	BEGIN
		Drop Table #CALCULATE_DIV
	END
		Select * INTO #CALCULATE_DIV
		FROM 
		(
			SELECT
				ELement_ID
			, TY_WND_DT
			, REGION
			, DISTRICT
			, DIVISION
			, VOL
			, ERR
			, FREQ
			, GOAL
			, EFF = Case When (FREQ/GOAL)>1 Then 1 Else (FREQ/GOAL) END
			--, EFF = CAST(FREQ/GOAL as numeric (25,2))
			, TIER
			, POSSIBLE_POINTS
			, ACTUAL_POINTS = CASE 
								WHEN @ELEMENT_TIER = 98 
								THEN ( [DADH1001].[src].[fn_Tier_98] ((FREQ/GOAL),@Element_Points))
							  
								WHEN @ELEMENT_TIER = 90 
								THEN ( [DADH1001].[src].[fn_Tier_90] ((FREQ/GOAL),@Element_Points))

								WHEN @ELEMENT_TIER = 80 
								THEN ( [DADH1001].[src].[fn_Tier_80] ((FREQ/GOAL),@Element_Points))

								else 0 end
		
			FROM #CALCULATE_B ) CALC_DIV
------------------------------------------------------------------------------------------
-- PUSH Final Data to Division Results List
-------------------------------------------------------------------------------------------

UPDATE DIVISION_RESULTS
	 	
SET    		  DIVISION_RESULTS.ID_1_Element_ID			= DIV.ELement_ID
		, DIVISION_RESULTS.ID_1_VOLUME				= DIV.VOL
		, DIVISION_RESULTS.ID_1_ERRORS				= DIV.ERR
		, DIVISION_RESULTS.ID_1_Freq				= DIV.Freq
		, DIVISION_RESULTS.ID_1_GOAL				= DIV.Goal
		, DIVISION_RESULTS.ID_1_Eff				= DIV.Eff
		, DIVISION_RESULTS.ID_1_Points				= DIV.Actual_Points
		, DIVISION_RESULTS.ID_1_Possible_Points	= DIV.Possible_Points
			 
		 
	FROM #CALCULATE_DIV DIV
	INNER JOIN [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIV] DIVISION_RESULTS
	ON
	DIV.TY_WND_DT		= DIVISION_RESULTS.TY_WND_DT 
	AND 
	DIV.REGION		= DIVISION_RESULTS.REGION 
	AND
	DIV.DISTRICT		= DIVISION_RESULTS.DISTRICT
	AND
	DIV.DIVISION		= DIVISION_RESULTS.Division

		
----------------------------------------------------------------------------------------------------------------------
-- This is pulling the RAW weekly generated data and using it to constrict a District Level Monthly Number to append
----------------------------------------------------------------------------------------------------------------------
If Object_ID('tempdb..#TEMP_M_RAW') is not null 
BEGIN
	Drop Table #TEMP_M_RAW
END

Select * into #TEMP_M_RAW
  
FROM 

(
	Select
		  Element_ID = @Element_ID
		, [MO_NUM]
		, [REGION]
		, [District]
		, sum([ID_1_Volume]) as Vol
		, sum([ID_1_Errors]) as Err
		, FREQ = case when CAST(SUM(ID_1_ERRORS)as float)=0 then cast(SUM(ID_1_VOLUME)as float)
	                 else cast(SUM(ID_1_VOLUME)as float)/cast(SUM(ID_1_ERRORS)as float) end
	  	, GOAL = @Element_Goal
		, TIER = @Element_Tier
		, POSSIBLE_POINTS = @Element_Points
  
	FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIS] MATRIX
  
	Group by MO_NUM, REGION, DISTRICT, ID_1_ELEMENT_ID
		   
) Temp_M

If Object_ID('tempdb..#CALCULATE_DIS_M') is not null 
BEGIN
	Drop Table #CALCULATE_DIS_M
END

Select * INTO #CALCULATE_DIS_M
FROM 
(
Select 

		Element_ID		      
 		, TY_WND_DT = '3333-01-'+ Mo_Num
		, MO_NUM
		, REGION
		, District
		, Vol
		, Err
		, FREQ 
	  	, GOAL 
		, Eff=Case When (FREQ/GOAL)>1 Then 1 Else (FREQ/GOAL) END
		, TIER 
		, POSSIBLE_POINTS = @Element_Points
		, ACTUAL_POINTS = CASE 
						WHEN @ELEMENT_TIER = 98 
						THEN ( [DADH1001].[src].[fn_Tier_98] ((FREQ/GOAL),@Element_Points))
							  
						WHEN @ELEMENT_TIER = 90 
						THEN ( [DADH1001].[src].[fn_Tier_90] ((FREQ/GOAL),@Element_Points))

						WHEN @ELEMENT_TIER = 80 
						THEN ( [DADH1001].[src].[fn_Tier_80] ((FREQ/GOAL),@Element_Points))

						else 0 end

		FROM #TEMP_M_RAW
) DIS_M

-------------------------------------------------------------------------------------------
-- PUSH Final Data to District Results List for MONTHLY DISTRICT
-------------------------------------------------------------------------------------------

UPDATE DISTRICT_RESULTS
	 	
SET  
		DISTRICT_RESULTS.ID_1_Element_ID		    	= DIS.Element_id
		, DISTRICT_RESULTS.ID_1_VOLUME				= DIS.VOL
		, DISTRICT_RESULTS.ID_1_ERRORS				= DIS.ERR
		, DISTRICT_RESULTS.ID_1_Freq				= DIS.Freq
		, DISTRICT_RESULTS.ID_1_GOAL				= DIS.Goal
		, DISTRICT_RESULTS.ID_1_Eff				= DIS.Eff
		, DISTRICT_RESULTS.ID_1_Points				= DIS.Actual_Points
		, DISTRICT_RESULTS.ID_1_Possible_Points	= DIS.POSSIBLE_POINTS
			 
		 
	FROM #CALCULATE_DIS_M DIS
	INNER JOIN [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIS] DISTRICT_RESULTS
	ON
	DIS.TY_WND_DT		= DISTRICT_RESULTS.TY_WND_DT 
	AND 
	DIS.REGION			= DISTRICT_RESULTS.REGION 
	AND
	DIS.DISTRICT		= DISTRICT_RESULTS.DISTRICT


---------------------------------------------------------------------------------------------------------------------
-- This is pulling the RAW weekly generated data and using it to constrict a Division Level Monthly Number to append
----------------------------------------------------------------------------------------------------------------------

If Object_ID('tempdb..#TEMP_MDIV_RAW') is not null 
BEGIN
	Drop Table #TEMP_MDIV_RAW
END

Select * into #TEMP_MDIV_RAW
  
FROM 

(
	Select
		  Element_ID = @Element_ID
		, [MO_NUM]
		, [REGION]
		, [District]
		, DIVISION
		, sum([ID_1_Volume]) as Vol
		, sum([ID_1_Errors]) as Err
		, FREQ = case when CAST(SUM(ID_1_ERRORS)as float)=0 then cast(SUM(ID_1_VOLUME)as float) else cast(SUM(ID_1_VOLUME)as float)/cast(SUM(ID_1_ERRORS)as float) end
	  	, GOAL = @Element_Goal
		, TIER = @Element_Tier
		, POSSIBLE_POINTS = @Element_Points
  
	FROM [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIV] MATRIX
  
	Group by MO_NUM, REGION, DISTRICT, DIvision, ID_1_ELEMENT_ID
 
) Temp_MDiv



If Object_ID('tempdb..#CALCULATE_DIV_M') is not null 
BEGIN
	Drop Table #CALCULATE_DIV_M
END


Select * INTO #CALCULATE_DIV_M
FROM 
(
Select 

		Element_ID		      
 		, TY_WND_DT = '3333-01-'+ Mo_Num
		, MO_NUM
		, REGION
		, District
		, Division
		, Vol
		, Err
		, FREQ 
	  	, GOAL 
		, Eff=Case When (FREQ/GOAL)>1 Then 1 Else (FREQ/GOAL) END
		, TIER 
		, POSSIBLE_POINTS = @Element_Points
		, ACTUAL_POINTS = CASE 
						WHEN @ELEMENT_TIER = 98 
						THEN ( [DADH1001].[src].[fn_Tier_98] ((FREQ/GOAL),@Element_Points))
							  
						WHEN @ELEMENT_TIER = 90 
						THEN ( [DADH1001].[src].[fn_Tier_90] ((FREQ/GOAL),@Element_Points))

						WHEN @ELEMENT_TIER = 80 
						THEN ( [DADH1001].[src].[fn_Tier_80] ((FREQ/GOAL),@Element_Points))

						else 0 end

		FROM #TEMP_MDIV_RAW
) DIV_M


-------------------------------------------------------------------------------------------
-- PUSH Final Data to District Results List for MONTHLY DISTRICT
-------------------------------------------------------------------------------------------

UPDATE DIVISION_RESULTS
	 	
SET  
		  DIVISION_RESULTS.ID_1_Element_ID		    	= DIV.Element_id
		, DIVISION_RESULTS.ID_1_VOLUME				= DIV.VOL
		, DIVISION_RESULTS.ID_1_ERRORS				= DIV.ERR
		, DIVISION_RESULTS.ID_1_Freq				= DIV.Freq
		, DIVISION_RESULTS.ID_1_GOAL				= DIV.Goal
		, DIVISION_RESULTS.ID_1_Eff				= DIV.Eff
		, DIVISION_RESULTS.ID_1_Points				= DIV.Actual_Points
		, DIVISION_RESULTS.ID_1_Possible_Points	= DIV.POSSIBLE_POINTS
			 
		 
	FROM #CALCULATE_DIV_M DIV
	INNER JOIN [DADH1001].[rpt].[t_82292_BSC_ESTIMATOR_MATRIX_RESULT_DIV] DIVISION_RESULTS
	ON
	DIV.TY_WND_DT		= DIVISION_RESULTS.TY_WND_DT 
	AND 
	DIV.REGION		= DIVISION_RESULTS.REGION 
	AND
	DIV.DISTRICT		= DIVISION_RESULTS.DISTRICT
		 	AND
	DIV.DIVISION		= DIVISION_RESULTS.DIVISION

	 
END
~~~~
----
We modify this baseline sproc with edits across the raw data sources and inject the same way to the three primary matrix tables for Raw, Division, and District.  We are also aggregating Monthy results as such as well.
All of this data then correlated in one final sproc that contains every single breakout and index calculation -- this is all done server side so that any excel refresh requirments are intantaneous and we aren't 
doing calculations within excel slowing down user response times


## Excel Configuration

Imported data comes into excel completely configured and ready to use as seen here:


![BSC_1](https://github.com/DonChart/Balanced_Scorecard_Estimator_Project/assets/168656623/b4de5a05-db55-488e-808f-b9df4b82d156)


What we do now is some excel trickery using pivot tables and showing/hiding rows using VBA to contain information on one screen for the end user and a button to jump from Weekly/Monthly and   
with the gridlines and row/column headers hidden the transition to tue user is seamless -- to indicate the changes I'll leave the Headins on here  

notice - a weekly presentation has rows 21-42 visible and a monthly 60-51 - 

###Weekly




 

 
 
  

  
  	


