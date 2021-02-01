#Amabilis v1.0

Amabilis is a stand-alone executable program for use by forestry professionals
and small forest landowners alike. Amabilis is able to virtually cruise twenty-two
different species of trees on the west coast of the US. These species hold commercial 
significance to the timber industry.


Amabilis reads minimal plot data to generate current stand conditions and virtually cruises
the trees for log grades and volume. Amabilis uses minimum top DIB (Diameter inside Bark) and 
minimum log length standards for optimal log grading according to the 
"Official Rules for the Log Scaling and Grading Bureaus". It should be noted that, as Amabilis 
cannot account for log defect or knot density, these grades are the "optimal grades' and may
not reflect the true conditions of the logs.


Amabilis requires six data and one optional datum. The six required data are...

    1) Stand Name-	Can be any meaningful string of characters for example: 'EX1'.

    2) Plot Number-	The plot number within the stands, needs to be an integer.

    3) Tree Number-	The tree number within the plot, needs to be an integer.

    4) Species- 	A two character string representing the Species Codes available, for example
        'DF' for Douglas-fir. Available species and codes can be found in the Species
        drop-down menu within the 'Data Sheet Tips' area of Amabilis.

    5) DBH-		Diameter at Breast Height. This needs to be any number greater than zero. Note Amabilis 
        does not cruise trees with less than 6.0 inches DBH, but will keep them in the count 
        for stand inventory.

    6) Plot Factor- For Variable Radius plots enter the Basal Area Factor (BAF).
        For Fixed Radius plots enter the negative-inverse of the plot area, for example a
        1/30th acre plot is -30.

    The optional datum is Total Height of the tree, however this is only optional after one Total Height is 
    recorded for the stand. Total Height data are used to get an average Height to Diameter Ratio (HDR) for 
    the stand, which fills in calculated heights for trees with missing height data.



The data input side of Amabilis starts in one of two ways...

	1)  	The user can manually enter data into the built-in formatted spreadsheet.
		This can be done by navigating to the 'File' menu and selecting 'New Sheet'.

	2) 	The user can load in a pre-filled CSV file. 
		This can be done by navigating to the 'File' menu and selecting 'Browse for CSV...'
			*This CSV file should be formatted correctly*
			-To generate a blank template CSV, navigate to the 'CSV' menu and select
			'Save New CSV...'


After completion of the data entry the user has a couple options...

	1)	The user can save the data to a new CSV file.
		This can be done by navigating to the 'CSV' menu and selecting 'Save New CSV...'.

	2)	The user can append to the current CSV file.
		This can be done by navigating to the 'CSV menu and selecting 'Append Current CSV'.

	3)	The user can append to another CSV file not currently being worked on in Amabilis.
		This can be done by navigating to the 'CSV' menu and selecting 'Append Other CSV...'.


Finally the user can use the data processing tools of Amabilis. To do this navigate to the menu
'Process Data' and select 'Go To Processing'. This will open a new window containing options 
for the tools available.

Before saving or appending CSVs, or opening the Data Processing window. Amabilis will check the data for errors.
These errors could be 'at least one height needed per stand', 'missing required data' or 'value errors' 
according to data type. If there are errors, Amabilis will return you to the data entry window and highlight the 
cells with errors.


The Data Processing window contains three main tools View Report, Print Report, and exporting to FVS databases.

	  1)	View Report and Print Report
		
      -To view or print the report, click the button with the stand's name you'd like to examine, OR click 
      the button 'All Stands' to generate a report for all stands within the CSV.

      -Enter the Preferred Log Length for the stand, in forestry this number is typically 32 or 40 feet.

      -Enter the Minimum Log Length for the stand, in forestry this number is typcially 8 feet or 16 feet, but 
      this can be set to zero for maximum scaling.

      -Click the 'View Report' button to view the report in-app or click the 'Print Report' button to export the 
      report to a PDF file.

	  2) 	Exporting to FVS Databases.
		
		  FVS is a program created by the US Forest Service and hosts a suite of tools for manipulating and projecting
      forest data. There are two main versions of FVS - FVS Legacy and FVS Modern. FVS Legacy requires a 
      locations file 'Suppose.loc' and a Microsoft Access database. FVS Modern can take databases in three formats, 
      Access, SQLite3 (.db) and Excel. Amabilis will export FVS formatted databases in all three types.

      To get started in exporting to the databases navigate to the 'Export to FVS' menu and select which 
      database to export to, or select 'Export to Multiple databases' to export multiple.

      For each stand in the user's CSV file, there are eight required data and two optional data to input.
      The required data are...
      1) Variant-		The variant locale for which FVS produces its models.
      2) Forest-		Each Variant includes Forest Codes of the forests within the Forest Service.
      3) Region-		The Forest Codes are within one of the ten regions of the Forest Service.
      4) Inventory Year-	The year that the stand inventory was taken.	
      5) Age-			The approximate age of the stand in years.
      6) Elevation-		The approximate or average elevation of the stand in feet above sea level.
      7) Site Species-	The Species Code for the species used in the site index calculation.
      8) Site Index-		The site index for the majority of the stand.

      The optional data are...
      1) Latitude		Example - 45.6018
      2) Longitude		Example - -123.4371
      These can be left blank, to which Amabilis has built-in lat/longs for the geographic centers of each region.

      When the button in the bottom left changes from 'Next Stand' to 'Export to Database' all stands have been completed
      and are ready to be exported to databases for use in FVS.





Further versions of Amabilis...
Likely version 2.0 will be a major update to include the ability to process full cruise data in addition to
the quick inventory data available currently.


Amabilis is produced by Pipeline Forestry of Washington State.
To contact Pipeline.

    Email - zachbeebe@pipelineforestry.com
    Phone - 425-931-8214
    Website - pipelineforestry.com


For tree data calculations, Amabilis uses a PYPI package also created by Pipeline Forestry, called timberscale.
For more information on timberscale goto...

https://github.com/zacharybeebe/timberscale

OR

https://pypi.org/project/timberscale/

		


