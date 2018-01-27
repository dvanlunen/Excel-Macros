# Excel Macros

Excel is the most commonly used tool for viewing data. Though most heavy lifting of data analysis happens elsewhere, often results are exported to xlsx so they are universal. Export options in other software are usually limited in terms of formatting, but formatting is essential to convey data. This repo is meant to list Excel macros that quickly format or manipulate spreadsheets so that they can be quickly shared.

Consider a starting sheet that is unformatted.

![Starting Table](/images/Before.png)


## TableFormat
This automatically sets up the sheet. Start in the top left of your table. Then it will emphasize the header of the table, automatically adjust column widths, format borders within the table, change the zoom of the sheet, and add rows above the table (where you can add a title to describe the table contents) among other edits.

If we apply this to the table above we get:

![TableFormat](/images/TableFormat.png)


## BarsUnder
This macro lets you specify a number of columns. Then it checks every row to see if the value of any of the first columns (where it checks the number specified) changes. If it does, it adds a red bar across the table. This is helpful when your data is in groups and you want to separate them visually. I use red because it is high contrast, but the macro can be set to have a more neutral color if that is preferred (sometimes I like green because it isn't as imposing).

To use this start in the leftmost column of the table in the first row of data (not the header). Running the macro on the starting table and specifying 2 columns, the output looks like this:

![BarsUnder](/images/BarsUnder.png)


## AlternateRowShade
This macro is very similar to BarsUnder. Instead of putting a bar under each group, it adds a background cell color for every other group so that you get alternating highlights. It can be used in combination with BarsUnder to separate groups at two different levels.

To use this start in the leftmost column of the table in the first row of data (not the header). Running the macro on the starting table and specifying 2 columns, the output looks like this:

![AlternateRowShade](/images/AlternateRowShade.png)


## WorksheetLoop
This macro applies the TableFormat macro to every worksheet in the excel workbook. I often use this if I output some result from R for each year into a different sheet, but want to format all the sheets.

The best part about this is that it can be customized depending on the need. Just change the code in the loop to fit.


## NotesSources
This macro adds a formatted list for you to add notes and sources to your analysis. When you pass your analysis along, you want people to know exactly what you did so you should write some notes. You also want them to know all the data you used so you can list the sources. This formatting is helpful because the numbers reference each other so you can quickly add new notes/sources and shift them around as well without too much manual work.

To use this macro select a cell below your table. If you select a cell in column A another column will automatically be added to the left to add the numbers. Also, note that you should be careful not to run this while selecting something above data because the numbers it throws in will overwrite cell values. If I click in cell A11 for the starting sheet and run this, the output looks like:

![NotesSources](/images/NotesSources.png)


## CommaFormat and CurrencyFormat
These are simple formats to round numbers to the nearest whole number and have commas where necessary. I find it helpful to have a single button to do this.
