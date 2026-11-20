PYTHONSCRIPTS/CALCULATORS

-BurningCalculator Builds BurningTotals.csv out of existing billable net tons/coil transfer logs. Data manually updated into Production.xlsx following format, and defined onto map.

-BreakingCalculator Builds BreakingTotals.csv out of existing unprepared scrap transfer logs. Data manually updated into Production.xlsx following format, and defined onto map.

-Loading Calculator LoadedBaskets report ran daily for consumption data, daily tons loaded into baskets, (intended to work on the other csv present in the folder with .py)
Consumption3 used once monthly for AverageConsumption.csv data that populates and is used in depleting past due piles(TOP RIGHT ## MANUALLY SET FOR DAYS TO DIVIDE DATA BY)
ENSURE SCRIPT IS RAN ON FULL MONTHS DATA EACH TIME SO EXPANDED HEAT VIEW HAS CORRECT DATA
MANUALLY ADD REMAINDER DATA FROM HEAT SEQUENCES DONE IN PREVIOUS MONTH, THEN TRANSPORT DATA TO CONSUMPTION4 TAB
Ensure you overwrite most recent date in Consumption1 Tab with new report data(Deduplicates Completed Heats)

-Receiving Calculator
Only works on CSV inside folder named exactly "Receiving", account for this

-Railcar Calculator
Python script for translating ASNs#-#.csv into completed XLSX, Supplier labels.

-------------------------------------------------------------------------

TOTAL YARD INVENTORY
Plus/Minus readout for overall inventory change in month
Debugging tooltip renders overall change at the moment.
NetInvChange metric derived from this math on final Receiving/Consumption values for month data

Stopped Pile Deviations -- Red Highlight on truck receipt     (GOAL TO SAVE/POPULATE THIS DATA FOR THE MONTH)
(,"Stopped":"yes") -- Only apply to one virtual pile, don't apply to both main/remote
This will also apply to any receipt into a pile identified as past due (greater than 6 months since last zero)
-Export this daily with a script to build stopped pile deviations report again

-------------------------------------------------------------------------

Jumping between historical dashboard/yardmap can cause "stickiness" tied to how window scales, any way to make this jump between dashboards more solid

More customization to XSXL Output from past due panel (conditional formatting, ages, days to deplete math)

Keyboard shortcuts to pull up Burning/Breaking/Bucketloading/Receiving/Railcars/Search panel

Red highlighting in history.xlsx transports partial heat sequence from prior month to first day of new month, labels it's day as the first of the month
THIS NOW HAPPENS AUTOMATICALLY -- PARTIAL BUCKET DATA FROM PREVIOUS YEAR IS MANUALLY APPENDED TO JANUARY DATA
JAVASCRIPT AUTOMATICALLY DEDUPES HEATS AND DISPLAYS ALL HEAT DATA ACROSS DAYCHANGES WITHIN THEIR OWN HEAT ON THE DAY THEY ARE COMPLETED

automatic detection of tickets that require void/reissue (punctuation in remark as example, data offset in row)

Translating "Heat_Type" Column to show "User Friendly" heats grades
Straightforward translation except for Nickel heats, which are 409###

LESS SPECIFIC POPOUT PANEL DATA BLOCKS WHEN EMPTY/TRANSITIONING TO NEW MONTH. "NO DATA" whenever determined should always show blank summary panel with the month selector functionality available to review historical data

Breaking pit "0 transfers this month" while no data is available

RED OUTLINE ON MARKER ICON.PNG IF THE PILE THAT EXISTS AT THAT MARKER IS NEGATIVE OR UNDER THE AVERAGE USE AMOUNT (<5000?)

Include all possible cutting/unprep piles in markers.json, Write javascript to not display pile in list if "last zero date" data is empty/missing
---------
"Consumed" metric in dashboard shows breakdown of materials consumed on each date with expanding rows.
"Heats" metric in dashboard shows breakdown of each instance of materials used during each "heat sequence" in expanding rows -- Drill down method to reach, click MoM month chart entry to open daily break down of heats, clicking any single day bar chart entry expands that days heat viewer
expanded heat viewer within dashboard shows "lot #" column in addition to the material translated from that lot #, no need to display the "lot #" column at all. (ACTUALLY MAKE MATERIAL LOT # SHOW UP AS HOVERTEXT ON MOUSEOVER MATERIAL NAMES)

Expanded truck viewer within dashboard needs more design work

More thought to how to display railcar data, all railcar "release" dates are actually included in History and Production.xlsx in the "Receiving data"
Railcars are currently ignored by the display javascript but can be seen within the Receiving tab to apply true final datestamps for railcars in our yard, and then the weight received from that car would only apply accurately to the railcar weight received per month. 
"# of railcars" released per day is an irrelevant metric for the type of production we track but the true release date is valuable for the math that is done to determine total weight received per month.

Click to copy feature on Scale gross on railcar popout panel should highlight red if the discrepancy showing is greater than 2500 pounds
