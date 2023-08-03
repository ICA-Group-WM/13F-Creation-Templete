# 13F-Creation-Templete
* Note: This is a file used for July/August of 2023, you will need to change some of the items in the code so it fits the files you're using \
* Second Note: There are some hard coded elements because of how the files were given and the results that were needed. If your file has different values for these, **make sure to change them!**
1) Download the python file and put it in the folder that has your LPL and Schwab Excel files in it
2) At the top of the file, there is the name of the files, change them to fit what works for you
  - Keep in mind, this is using one file as a xlsx and one as a csv. If this isn't the case for you, you might need to change some of the code so it works for you!
3) The output file HAS to stay xlsx initally, you will get data loss if it is a csv file as an output. This is a side affect of using pandas and python.
4) The XML file output is using xml tree syntax for the output. This is standard use for the SEC and can be found here: https://www.sec.gov/page/edgar-how-do-i-create-xml-information-table-form-13f-using-excel
5) Once you have everything changed and set up for you, when you run the file, the output files will be saved in the same folder you are working in.
