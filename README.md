# Automatic Excel Data Ingestion & Analysis
#### Using Microsoft Excel macros and formatting to create an automatic ingestion solution for a small bakery business
### Introduction
  This project aims to develop a fully functional data ingestion solution for a small business using Microsoft Excel. Businesses conduct 10s to hundreds of transactions a day and as such, having to update
all their tables manually can be quite tedious. Through this data ingestion solution, the Excel macros coded in will be able to do all the heavy lifting, making updating the table and adding new tables a seamless process. 

  On top of this imported dataset, using chatgpt another table will be created to track prices and price changes. Then we will create a third table using lookups and functions to track sales by month. This data will then 
allow us to perform some analysis and visualization using tools like pivottables and charts. The dataset being used is for a small bakery business; the link to the base dataset can be found [here](https://www.kaggle.com/datasets/akashdeepkuila/bakery?resource=download). 

### Step 1: Importing the Data & Formatting

  Now that the dataset is downloaded, we must import it into Excel. After using the 'Data' tab's 'Get External Data' - 'From Text', the csv file can be imported into the worksheet. Making sure to use commas as the delimiters, our dataset now looks like this:

![image_alt](https://github.com/brianhornick/Automatic-Excel-Data-Ingestion-and-Analysis/blob/main/Photos/Screenshot%202025-07-07%20150935.png?raw=true)

  Next, let's turn on the macro recorder and format our table. After formatting our header and fixing the column alignments, the table now looks much cleaner, and we have a recorded macro that can be used to format other tables we may want to add automatically. I also changed the name of the worksheet. Here is what the table now looks like:

![image_alt](https://github.com/brianhornick/Automatic-Excel-Data-Ingestion-and-Analysis/blob/main/Photos/Screenshot%202025-07-08%20140657.png?raw=true)

### Step 2: Creating and Implementing the Price Tracker Table

  
