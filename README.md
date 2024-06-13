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
