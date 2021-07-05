# stock-analysis

### **Overview** ### 
The purpose of this project was to create an Excel workbook that can analyze an entire dataset of stocks performance for a given year with buttons to initiate the code for a simple user interface. In order to analyze the performance of a stock and its annualized return, the data must include the ticker, the daily performance of the stock and the number of daily trades. This type of analysis will require a long list of information for the program to review in order to provide meaningful results. Depending on the quantity of stocks in the dataset, it could be thousands or hundreds of thousand of lines of data. In order to accomplish the goal of this analysis, a program capable of analyzing the correct areas of the data was created. Later that code was refactored, in other words, the design of the code was restructured in order to improve the speed of the analysis and the clarity of its operation. The outcome was a code capable of managing larger amounts of stocks data more efficiently. 

![Example_of_Data_Better](https://user-images.githubusercontent.com/85839235/124513381-aeb0d200-dda8-11eb-8076-3f482e312e96.png)
This is an example of the data analyzed. It contained over 3,000 lines of information to analyze for only 12 stocks. 

### **Stock Performance**
The original request for the construction of the code was to analyze the performance of the securities for Daqo New Energy “ticker: DQ” in order to evaluate it as an investment opportunity. The results of the analysis showed that “DQ” had a negative performance for the year 2018. It erased most of the gains from the year 2017. It was also traded about 3x more often during 2018 than on the previous year. While the trading volume metric by itself doesn’t provide sufficient data as a standalone metric, when paired with other pieces of information it may be something worthwhile to consider 1️⃣. Nonetheless the 3x increase on trading for the year 2018 and the loss of value of approximately 63% in value may be a reasonable call to research more about the company profitability and value before investing. With this in mind the code was expanded in order to analyze multiple stocks performance. 

###### **Stocks Performance Comparison**
![Performance_Comparison](https://user-images.githubusercontent.com/85839235/124512938-a1471800-dda7-11eb-8ae0-c1e6031270ec.png)


### **Overview of the Code**

#### **Original**

Excel was used in order to source the data and to present the results of the analysis. VBA, or Visual Basic for Applications, was the programming software used to create the application. The original code was designed to identify the correct Excel Year workbook the user wanted to review, gather the necessary information for every stock and analyze the entire dataset multiple times in loops in order to provide a summary of the stock’s performance including the number of times it was traded during that year and its annualized return. The program also allowed the end user to activate the analysis with the click of a bottom within Excel. The user would receive a table of data that included the stock performance along with color formatting depending in its performance for the year. 

###### **Visual of User Interface**
![Code_Example_1](https://user-images.githubusercontent.com/85839235/124513240-55e13980-dda8-11eb-89ca-2fcb464f4e3b.png)

###### **Visual of Data Output**
![Code_Example_2](https://user-images.githubusercontent.com/85839235/124513255-5f6aa180-dda8-11eb-955d-b49852b1e8b8.png)


#### **Refactored**
The original code performed the analysis and yield the results requested correctly. However, the method in which the code gathered the information, looping over all the data multiple times, was inefficient and would have required considerably longer amount of time in order to process larger datasets. In the refactoring process, the analysis of the information was kept the same but the process in which the information was taken from the dataset was improved. It allowed the code to gather the information requested from each row for all the stocks ticker and only needed to loop over the dataset once. This resulted in a code that was 5 times faster than the original but returned the same analysis and results. 

###### **Performance Comparison**
![2018_Analysis_Original](https://user-images.githubusercontent.com/85839235/124511851-11a06a00-dda5-11eb-95d6-40747dff0c7c.png)
The run time of the original code for the Year 2018 was approximately 0.71 seconds.

![2018_Analysis_Refactored](https://user-images.githubusercontent.com/85839235/124511886-20871c80-dda5-11eb-89fa-ccded11b8fd7.png)
The run time for the same analysis using the refactored code was approximately .14 seconds. 

### **Structure of the Code**
###### **Original**
![Original_Code_Sample](https://user-images.githubusercontent.com/85839235/124512063-870c3a80-dda5-11eb-9c86-dd53565053d6.png)
The original code was a simpler design. It reviewed the data looking for a specific value and once it was done it would move into the next value and loop again until it captured the information for every stock.

###### **Refactored**
![Refactored_Code_Sample](https://user-images.githubusercontent.com/85839235/124512092-955a5680-dda5-11eb-8bdd-abaee341f09f.png)
The refactored code was longer in its instructions but allowed VBA to gather the information quicker and more efficiently. 

### **Conclusion regarding Refactoring the Code for Stocks**
It is reasonable to conclude that the refactored code is faster and more effective in processing the analysis. It will also be able to handle bigger datasets better than the original. However, the construction of this specific refactored code took a considerably longer amount of time and skills in order to build. While the original code took minutes to design, the refactored code took multiple hours in order to streamline and troubleshoot. More arrays were necessary to include in the code and linking them to the correct indexes of the original code took a substantial amount of time. With that in mind, this code arguably benefited from having been refactored. It will allow the end user to analyze bigger data sets in the future quicker and without being dependent on the capabilities of the computer used to run the program. This is extremely important since the data that the end user may want to analyze is vast. As of 2020 there were close to 6,000 companies trades on the NYSE and Nasdaq, with over 11,000 securities on alternative trading systems2️⃣. 


### **Conclusion on the practice of refactoring**
The original design of a successful code will not always be the best way to achieve the result wanted or the most efficient path. Refactoring will arguably be a noble goal to pursue in the writing of codes and analysis but may not always be necessary. It would be wise to adjust the scope in designing a code depending on its intended purpose. For a small set of data or an analysis that will be used infrequently or even once, refactoring may not be necessary. It may be even detrimental for an Analyst to refactor or optimize every code if there is no gain or purpose for it. It may take valuable time or resources that may be better used in other tasks or parts of the analysis. Refactoring will be a good practice when writing code if the task calls for it. 

### Reference Links

[1️⃣] https://smartasset.com/financial-advisor/high-volume-stocks

[2️⃣] https://www.marketwatch.com/story/the-number-of-companies-publicly-traded-in-the-us-is-shrinkingor-is-it-2020-10-30?mod=investing
