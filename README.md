# VBA-Chalenge
Crowdfunding platforms like Kickstarter and Indiegogo have been growing in success and popularity since the late 2000s. From independent content creators to famous celebrities, more and more people are using crowdfunding to launch new products and generate buzz, but not every project has found success.

To receive funding, the project must meet or exceed an initial goal, so many organizations dedicate considerable resources looking through old projects in an attempt to discover “the trick” to finding success. For this Challenge, I will organize and analyze a database of 1,000 sample projects to uncover any hidden trends.
Creating a script in VBA that loops through all the stocks for one year.
# Dataset:
[CrowdfundingBook dataset](https://github.com/dilqvl62/VBA-Chalenge/blob/main/Data%26Output/CrowdfundingBook.xlsx)

# Objectives
**Step 1**
 * Use conditional formatting to fill each cell in the outcome column with a different color
   * Create a new columns
  
![Screen Shot 2023-10-23 at 12 57 35 PM](https://github.com/dilqvl62/VBA-Chalenge/assets/107519883/a189929f-8262-474f-874b-b33060fe0a91)

  * Create a new sheet with a pivot table that analyzes the initial worksheet to count how many campaigns were successful, failed, canceled, or are 
    currently live per category **and** a stacked-column pivot chart that can be filtered by country based on the table created .

  ![Screen Shot 2023-10-23 at 12 59 44 PM](https://github.com/dilqvl62/VBA-Chalenge/assets/107519883/6a8d0479-d93f-4fd1-a605-668092ff3aca)

* Create a new sheet with a pivot table that analyzes the initial sheet to count how many campaigns were successful, failed, or canceled, or are currently live per sub-category **and** a stacked-column pivot chart that can be filtered by country and parent category based on the table created.
  
![Screen Shot 2023-10-23 at 1 00 47 PM](https://github.com/dilqvl62/VBA-Chalenge/assets/107519883/8609a68c-2fdf-41fa-9945-1d2b294f31ff)

*convert Launched_at and deadline column into Excel date format*

* Create a new sheet with a pivot table **and** a pivot-chart line graph that visualizes this new table.

![image](https://github.com/dilqvl62/VBA-Chalenge/assets/107519883/d7ffca3e-7c37-4b6f-9098-cf3b266496eb)

# Crowfunding Goal Analysis

* Create a new sheet with 8 columns:
   * Create a line chart that graphs the relationship between a goal amount and its chances of success, failure, or cancellation.

![image](https://github.com/dilqvl62/VBA-Chalenge/assets/107519883/d96d545a-fb11-4da6-9e1a-3c22a202c292)
