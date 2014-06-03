ft-data
=======

# Background #
This repository was created as I heard someone talking about trying to import data from the Financial Times data into an Excel spreadsheet.

The data was in the form of a pdf that was downloaded from the website, but they couldn't cut and paste the data from the pdf into Excel.

As anyone that has experienced cut and paste from a pdf, whenever they tried that, the formatting would be all messed up. Rows were out of order and looked to be random and all the data stuffed into the first cell.

At first I suggested looking for a pdf-to-excel converter.

I found a couple online and tried them, but the results were rather disappointing. Some data disappeared during the conversion and the resulting MS Excel workbook wasn't much use as the formatting was still messed up.

So I decided that I can do better and wrote a php script to format the data into something useful to the end user. All the relevant data should be split into rows and columns just like the original pdf document.

After doing this, I thought it might be better if I could do the same thing but as an MS Excel macro. So the user could just cut and paste the data into MS Excel and press a button to format it.

Writing the macro was a bit painful as it was a long time since I had written any VBScript. I'd also forgotten how annoying and ugly it was to write VBScript! Progress was slow and google was consulted a lot.

But I did end up with a workbook containing a button that could be pressed to format the data into something usable.

# Data #

The data pdf was "FT Guide to World Currencies" from http://markets.ft.com/research/Markets/Currencies

You can download the pdf at the bottom of that page by selecting "Currencies" and "FT Guide to World Currencies", choosing a date and clicking "download"

An example pdf is at this url http://markets.ft.com/RESEARCH/markets/DataArchiveFetchReport?Category=CU&Type=WORL&Date=05/30/2014


# Usage #
First you need to visit the ft.com website and get the pdf.

Next you need to cut and paste the data into a MS Excel or a spreadsheet program and save it as a csv. What you should end up wih is something like FT.csv

## PHP Script ##
The PHP script can be run over the command line and requires the path to the csv file as the first argument
`php ftCurrencies.php path/to/csv

This will read in the csv and process it, outputting a new csv file where the data is formatted and sorted by Country.

The output file should be able to be read by MS Excel or any spreadsheet program.

## Excel Workbook ##
Again visit the ft.com website and getthe pdf.

Next you need to cut and paste the data into the MS Excel Workbook on Sheet 1, starting at A2.

Once that is done, press the "Format Currencies" button on the "unformatted" worksheet.

This should go through all the rows, format the rows and output them onto "formatted" worksheet.

Rows that have no data should be skipped and be empty in the "formatted" worksheet.

Country and Currency have been split into separate columns as this was a request by the user.


# Conclusions #

* Cutting and pasting from pdf documents is a bit crap.
* Writing VBScript is not something I enjoy. It's rather ugly and trying to do something relatively simple takes a long time.
* Could be a useful exercise to do with different programming languages to get more familiar with them when I'm looking for something to do.
