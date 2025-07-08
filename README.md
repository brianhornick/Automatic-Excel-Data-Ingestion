# Automatic Excel Data Ingestion & Analysis
#### Using Microsoft Excel macros and formatting to create an automatic ingestion solution for a small bakery business
### Introduction
  This project aims to develop a fully functional data ingestion solution for a small business using Microsoft Excel. Businesses conduct 10s to hundreds of transactions a day and as such, having to update
their transactions table manually can be quite tedious. Through this data ingestion solution, the Excel macros coded in will be able to do all the heavy lifting, making updating the table a seamless process. 

We will also do some analysis on this data, using tools like pivottables, functions, and lookups. The dataset being used is for a small bakery business; the link to the dataset can be found [here](https://www.kaggle.com/datasets/akashdeepkuila/bakery?resource=download). 

### Step 1: Importing the Data & Formatting

  Now that the dataset is downloaded, we must import it into Excel. After using the 'Data' tab's 'Get External Data' - 'From Text', the csv file can be imported into the worksheet. Making sure to use commas as the delimiters, our dataset now looks like this:

![image_alt](https://github.com/brianhornick/Automatic-Excel-Data-Ingestion/blob/main/Screenshot%202025-07-07%20150935.png?raw=true)

  Next, let's turn on the macro recorder and format our table. After formatting our header and fixing the column alignments, the table now looks much cleaner, and we have a recorded macro that can be used to format other worksheets and tables automatically. I also changed the name of the worksheet:
  
