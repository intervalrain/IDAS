# About IDAS+
**Integrated Data Analaysis System Plus**
For data analysis with summary table and plot charts.

# Manual
## WorkSheet
+ SPEC
  + SPEC setting for parameters, including Device types, Facility(normailization), Unit, Target, FF, SS.
+ SPEC_List
  + SummaryTable item setting.
+ Corner
  + "getCorner" function setting, which is correllated with "genChartSheet" and "updateChartSetting".
+ Grouping
  + Set grouping with selected wafers.
  + Remember to follow the naming rules!!
+ Formula
  + Mismatch formula setting.
+ ChartType
  + Generate Charts and genWaferMapping item setting.
+ RECEIVER
  + Auto-rpt and mailing setting.
+ PPT
  + genPPT format setting.
+ PlotSetup
  + Chart format setting for Generate Charts.

## Command Bar
+ Load Data
  + Load *.rpt data to start analyzing.
  + Now, multiple files selection is available. (Prevent different WAT recipe)
+ Initial
  + If SPEC or data is revised manually, click Initial button to initialize the system setting.
+ Select Wafer
  + To select specific wafer to anaylize
+ Summary Table
  + To output summary table with format setup before.
+ Generate Charts
  + To output charts with format setup before.
+ Generate Charts
  + To use extra advanced functions.
## Manual Functions
+ File Class
  + Export File
    + To tranform the IDAS format back into raw data format, and export to selected directory.
  + Content Query
    + Select multiple *.rpt files to show simple info and contents of each files.
  + Load Spec File
    + To move worksheets from other Excel, usually used for IDAS version updation.
  + Load Long File
    + Load Data with *.long file. (single file only)
  + UEDA to IDAS
    + To tranform UEDA format to IDAS.
  + Rows to Table
    + To quickly generate table from rows data. (e.q. Model pdf)
+ Mismatch Class
  + Gen Mismatch
    + Generate mismatch report.
  + Plot Mismatch Chart
    + Plot mismatch charts by grouping
+ Chart Class
  + Gen Chart Sheet
    + To generate chart setting sheets with GUI.
  + Gen Wafer Map
    + To display data distribution of selected parameter.
  + Renew All_Chart
    + To renew ALL_Chart sheet after revision of Charts.
  + Gen PPT
    + To generate PPT with setting "PPT" sheet.
  + Pin Scatter
    + To pin All chart sheets, to prevent overwriting while generating charts.
  + Recount Corner Rate
    + To recount corner rate after removing out-liner points.
  + Generate Single Chart
    + To generate single chart with selected chart setting. (referring to selected cell)
+ Other
  + Get Coordinate
    + To generate coordinates by *.waf on the WAT system. (Note: Only used for there's coordinates missing)
  + Hint
    + To show some hints to use IDAS+.

# Copyright 
@author: [Rain Hu](https://intervalrain.github.io/posts/aboutme/)  
@email: [intervalrain@gmail.com](intervalrain@gmail.com)  
@github: [https://github.com/intervalrain](https://github.com/intervalrain)  
@website: [https://intervalrain.github.io](https://intervalrain.github.io)

# Version
Version Log:   
2022/05/08 Fix standard deviation's precision in "summaryTable".
2022/04/18 Fix "getCorner" function and simplify setting in "Corner" Sheet.  
2022/03/25 Fix ":" problem with ";" operator. Fix "UEDA to IDAS" function. Add "shrink" for "gen Mismatch" function.  
2022/01/15 Optimize pinScatter and UnpinScatter function. Add **genSingleChart** manual function.  
2022/01/06 Optimize unit setting for Diff.  
2021/12/30 Add alert if grouping setting fails while Summary Table or Generate Charts.  
2021/12/20 Add "Grouping" Sheet for grouping control, including "Summary Table" & "Generate Charts". Add "comp", "R2" keyword in "SPEC_LIST" to compare data with BSL(wafer_1). Add "Trend." and "Trend%" keyword to calculate Linear Regression by previous 2 rows. (x and y).  
2021/12/14 Add "SplitBy Group" function in Generate Charts.  
2021/12/02 Auto Corner plot setting.  
2021/11/12 Fix UEDA to IDAS ont lot problem.  
2021/10/27 New add Unpin Scatter, Optimize genChartSheet and genPPT.  
2021/10/22 Optimize genWafermap, now select_wafer is also available for genWafermap.  
2021/10/21 Add renew All_Chart, genPPT, pin Scatter. Optimize getSpec, genChartSheet(Thru-W/L auto getSpec), ContentQuery, Boxtrend chart plotting.  
2021/10/12 Fix genChartSheet problem and fix UEDA to IDAS problem.  
2021/10/06 Add WaferMapList to enable mulitple parameter wafer map ploting.  
2021/10/04 Optimize UEDA to IDAS. Use "Diff.OfParamter*n" in SPEC_List to mutliply value with n.  
2021/09/30 Optimize import UEDA function (Available to import inline CD and WAT data).  
2021/08/31 Add a char before wafer num when import UEDA in order to select wafers.  
2021/08/09 Add UEDA to IDAS function.  
2021/05/07 Add wafermap with spec_high and spec_lo.  
2021/05/05 Test Autorpt function.  
2021/04/17 Fix Auto-Highlight current direction in summary table.  
2021/04/15 Optimize Load Multi-Files function.(Auto-arrange by lot and wafer-order)  
2021/04/12 New add Load Multi-Files function.  
2021/03/24 New add Sensitivity plot chart function.  
2021/02/25 Highlight main leakage direction with bold font style for summary table.  
2021/01/11 New add Content query. (To list multi-selected files with brief info)  
2021/01/05 New add GenChartSheet function.  
2020/12/28 Mismatch summary optimize.  
2020/12/23 New add Mismatch summary and plot chart function.  
2020/11/26 Reconstruct and optimize SPEC_List functions. (Detail see SPEC_List Column)  
2020/11/20 Update site coordinate update function.  
2020/11/06 New add wafer mapping function.  
2020/10/25 Reconstruct IDAS.
