#BUSINESS DATA ANALYSIS


#Importing the necessary Python modules 
import pandas as pd 
import openpyxl 
import numpy as np 
import matplotlib.pyplot as plt


#Part One: Loading and Inspecting Excel Data 
 
#1. Loading and reading the file 'yelp.xlsx' 
#Accessing the file 
xl = pd.ExcelFile('yelp.xlsx')
#Loading each sheet onto a separate dataframe 
df_yelp = xl.parse('yelp_data')
df_cities = xl.parse('cities')
df_states = xl.parse('states')


#2. General inspections of the file 
#Inspecting the shape (rows x coloumns) of first sheet
shape = df_yelp.shape
print('Number of coloumns:', shape[1])
print('Number of rows:', shape[0])
print('')


#Inspecting the coloumn headers of the first sheet
print('Coloumn headers in the sheet \'yelp_data\':')
for column in list(df_yelp.columns):
    print(column)
print('')


#Display the first 5 enteries of data in the first sheet
print('The first 5 enteries off the \'yelp_data\' sheet:')
print(df_yelp.head())
print('')


#Part Two: Merging and Updating Data 
#1. Merging the three sheets, 'yelp_data', 'cities', and 'states', into one dataframe 'df'
#Merging the two dataframes 'df_yelp' and 'df_cities'
df = pd.merge(left=df_yelp, right=df_cities,           #specifying the dataframes to merge 
            how='inner',                               #specifying the type of data merging 
            left_on='city_id', right_on='id')        #specifying the common coloumns to merge on

#Merging the dataframes 'df' and 'df_states' 
df = pd.merge(left=df, right=df_states, how='inner', left_on='state_id', right_on='id')


#Displaying the first 5 rows after merging 
print('Dataframe after merging:')
print(df.head())
print('')


#2. Updating the dataframe 
#Deleting duplicate coloumns 
del df['id_x']
del df['id_y']


#Renaming the coloumns, 'category_0' and 'category_1', to be more representative of their data 
df.rename(columns={'category_0': 'business type',       #specifying the old (left) and new (right) coloumn titles
          'category_1': 'service type'}, 
          inplace=True)

#Displaying the coloumn titles again 
print('Coloumn titles after renaming:\n', list(df.columns))
print('')

#Rearranging the coloumns order 
rearranged_coloumns = ['name', 'business type', 'service type', 'stars', 'review_count', 'take_out', 'city', 'city_id', 'state', 'state_id']
df = df.reindex(columns=rearranged_coloumns)

#To preview the dataframe after updates
print('Updated dataframe:')
print(df.head())
print('')

#Displaying only the coloumns 'business type' and 'service type'
print('The following table displays the type of business for each business listed and the service they offer:')
print(df[['business type', 'service type']])
print('')


#Part Three: Querying Data 
#General queries 

#1. Report businesses located in Las Vegas only 
#Filtering data for Las Vegas (LA) enteries only 
LA_filter = df['city'] == 'Las Vegas'
df_LA = df[LA_filter]

#reporting businesses located in LA only 
print('The following table displays businesses located in Las Vegas only:')
print(df_LA)
print('')


#2. Report businesses classified as 'Restaurants' (in either categories, business/service type)
#Filtering data for 'Restaurants' in coloumns 'business type' and 'service type'  
rest_filter1 = df['business type'] == 'Restaurants'
rest_filter2 = df['service type'] == 'Restaurants'
df_restaurants = df[(rest_filter1 | rest_filter2)]         #i.e., restaurants is 'either' in first 'or' in second coloumn

#reporting businesses classified as 'Restaurants' only 
print('The following table displays all the restaurants featured in the data set:')
print(df_restaurants)
print('')


#3. Report businesses classified as either 'Restaurants' or 'Bars' (in both categories)
#Filtering data for both 'Restaurants' and 'Bars' simultaneously  
rest_filter = df['business type'].isin(['Restaurants', 'Bars'])
bars_filter = df['service type'].isin(['Restaurants', 'Bars'])
df_rest_bars = df[(rest_filter | bars_filter)]      

#reporting businesses after filtering 
print('The following table displays businesses classified as either restaurants or bars:')
print(df_rest_bars)
print('')


#4. Report Restaurants and Bars located in Las Vegas only 
#Adding a city filter to extract those in LA 
LA_filter = df['city'] == 'Las Vegas'
LA_rest_bars = df[LA_filter & (rest_filter | bars_filter)]       #i.e., the city filter AND either restaurants filter OR bars filter must be true

#reporting bars and/or restaurans in LA 
print('The following table lists businesses in Las Vegas that are classified as restaurants or bars:')
print(LA_rest_bars)
print('')


#More specific queries 
#5. How many Beauty and Spa centers are there in the city of Henderson? 
#Filtering cities for Henderson
Henderson_filter = df['city'] == 'Henderson'

#filtering businesses for Beauty & Spa centers
BnS_filter1 = df['business type'] == 'Beauty & Spas'
BnS_filter2 = df['service type'] == 'Beauty & Spas'

#creating a filtered dataframe 
df_BnS_Henderson = df[Henderson_filter & (BnS_filter1 | BnS_filter2)]

#Reporting the number of beauty & spa centers
print('The total number of Beauty and Spa centers in Henderson is:', len(df_BnS_Henderson))
print('')


#6. What type of service does the business 'Dairy Queen' offer? 
#Filtering business names for 'Dairy Queen'
DQ_filter = df['name'] == 'Dairy Queen'
df_DQ = df[DQ_filter]
print('Dairy Queen offers:', df_DQ['service type'].values[0])
print('')


#7. How many 'dive bars' are there in Las Vegas? 
# Can you recommend one in that city with at least a 4 star rating?

#filtering data by city 
LA_filter = df['city'] == 'Las Vegas' 

#filtering data for dive bars in business and service categories
divebars_filter1 = df['business type'] == 'Dive Bars'
divebars_filter2 = df['service type'] == 'Dive Bars'

#creating a filtered dataframe 
LA_divebars = df[LA_filter & (divebars_filter1 | divebars_filter2)]

#reporting the number of dive bars 
print('The total number of dive bars in Las Vegas is:', len(LA_divebars))
print('')

#7.2. Recommending a dive bar with at least a 4-star rating (randomly)
#adding a rating filter
rating_filter = LA_divebars['stars'] >= 4
LA_divebars_filtered = LA_divebars[rating_filter]

#resetting the index of the dataframe LA_divebars
LA_divebars_filtered.reset_index(drop=True, inplace=True)

#getting a random index to make a random recommendation
import random 
random_index = random.randint(0, len(LA_divebars_filtered)-1)

#reporting a random LA dive bar with 4-star rating or above 
random_divebar = LA_divebars_filtered.iloc[random_index]
print('Here\'s your recommendation of a highly rated dive bar in LA:', random_divebar['name'])
print('')


#8. Which city has the largest number of pizza shops?  
# Can you recommend a random one with at least a 4-star rating and which offers take-outs? 
#filtering data for pizza restaurants
pizza_filter1 = df['business type'].str.contains('Pizza')
pizza_filter2 = df['service type'].str.contains('Pizza')

#creating a dataframe with pizza businesses only
df_pizza = df[(pizza_filter1 | pizza_filter2)]

#Grouping and filtering data based on city
cities = df_pizza.groupby('city').groups         #returns a dictionary 
maxcount = None 
maxcity = None 
for city, val in cities.items(): 
    if maxcount is None or len(val) > maxcount:        #to get the city with most pizza businesses
        maxcount = len(val)
        maxcity = city 
    else:
        continue 

#Reporting the result 
print('The city with the most pizza shops is:', maxcity)
print('')

#8.2. Recommending a random pizza shop with a 4-star rating or above and which offers take-outs
#adding filters for each of the requirements
city_filter = df_pizza['city'] == maxcity
rating_filter = df_pizza['stars'] >= 4 
take_outs_filter = df_pizza['take_out'] == True

#creating a dataframe with data meeting the specified requirements 
df_pizza_filtered = df_pizza[city_filter & rating_filter & take_outs_filter]

#resetting the index of the dataframe 
df_pizza_filtered.reset_index(drop=True, inplace=True)

#selecting one at random 
random_indx = random.randint(0, len(df_pizza_filtered)-1)
random_pizza = df_pizza_filtered['name'].iloc[random_indx]

#reporting the results 
print('Here is your recommendation:', random_pizza)
print('')



#Part Four: Grouping, Summarizing, and Statistically Analyzing Data 
#1. Report the mean and standard deviation of rating scores of businesses for each city separately 
#grouping data by city and selecting statistics to perform
city_stats = df.groupby('city').agg([np.mean, np.std])['stars']

#reporting the rating mean and standard deviations per city
print('The following table displays the mean and standard deviation scores of businesses per city:')
print(city_stats)
print('')


#2. Report the mean star ratings and total number of reviews of restaurants for each city separately
#filtering dataframe for resturants only 
rest_filter1 = df['business type'] == 'Restaurants'
rest_filter2 = df['service type'] == 'Restaurants'
df_restaurants = df[(rest_filter1 | rest_filter2)]

#grouping data by city and selecting statistics for each coloumn to analyze
city_rest_stats = df_restaurants.groupby('city').agg({'stars': np.mean, 'review_count': np.sum})

#reporting the mean of star ratings and sum of reviews for restaurants per city 
print('The following table demonstrates the average rating score and total sum of reviews for restaurants per city:')
print(city_rest_stats)
print('')

#2.2. Based on the results, which city has the best restaurants?

#traversing through the dataframe to get the name of the city with the highest mean rating
maxrating = None 
maxcity = None 
for rating in city_rest_stats['stars']:
    if maxrating is None or rating > maxrating: 
        maxrating = rating 
        rating_filt = city_rest_stats['stars'] == rating 
        maxcity = city_rest_stats[rating_filt].index.values
    else:
        continue

#Reporting city whose restaurants received the highest mean rating
print('The city with the best restaurants:', maxcity[0])
print('')


#3. Create a pivot table showing the mean star rating and mean reviews 
# of businesses, grouping the results by city

pivot_tab = pd.pivot_table(df,              #specifying the dataframe to analyze
                index=['city'],             #specifying the coloumn to group data by 
                values=['stars', 'review_count'],         #specifying the coloumns to analyze 
                aggfunc=np.mean)            #specifying the statistic to perform 

#reporting the results of the table 
print('The following pivot table displays the average star rating and review count for businesses per city:')
print(pivot_tab) 
print('')


#4. Create a pivot table showing the mean star rating and total sum of reviews
# for businesses classified as 'Hotels & Travel', grouping the results by state and city

#filtering data for businesses classified as Hotels & Travel
Hotels_filter1 = df['business type'] == 'Hotels & Travel'
Hotels_filter2 = df['service type'] == 'Hotels & Travel'

#creating dataframe with hotels & travel businesses 
df_Hotels = df[(Hotels_filter1 | Hotels_filter2)]

#creating a pivot table to calculate their mean star ratings and sum of reviews 
pivot_Hotels = pd.pivot_table(df_Hotels, 
                            index=['state', 'city'],                #grouping results by state followed by city 
                            values=['stars', 'review_count'],         #specifying the coloumns to analyze
                            aggfunc={'stars': np.mean, 'review_count': np.sum})          #specifying the calculation to perform per coloumn

#reporting the pivot table
print('The following pivot table displays the average star ratings and total sum of reviews\nfor hotels & travel businesses, based on their state and city locations:')
print(pivot_Hotels)
print('')



#Part Five: Data Visualization 
#1. Create a histogram to compare the frequency distribution of rating scores 
# of businesses in Pittsburgh vs. Las Vegas

#filtering data for cities Pittsburgh and Las Vegas
Pitts_filter = df['city'] == 'Pittsburgh'
Pitts_rating = df[Pitts_filter]['stars']

LA_filter = df['city'] == 'Las Vegas'
LA_rating = df[LA_filter]['stars'] 


#plotting histogram to show frequency distribution of different businesses ratings in Pittsburgh
plt.hist(Pitts_rating,
         label='Pittsburgh',
         color='#9cd095',
         alpha=0.8,
         linewidth=1, edgecolor='k',
         bins='auto')

#plotting histogram to show frequency distribution of different businesses ratings in LA
plt.hist(LA_rating, 
        label='Las Vegas',             #setting the label/description of the histogram
        color='#95c6d0',              #specifying the color of histogram bars
        alpha=0.6,                   #setting the degree of bars transparency
        linewidth=1, edgecolor='k',     #setting the bar edges' width and color
        bins='auto'                 #divides the bins along the x-axis automatically
        )


#Adding a title to the histogram 
plt.title('Distribution of Rating Scores for Las Vegas and Pittsburgh Businesses')
#Adding labels to the histogram axes 
plt.xlabel('Rating Score')       
plt.ylabel('Frequency of Rating Score')
#Adding a legend to describe the histogram better 
plt.legend(title='Business location:', loc='best')      #specifies the title and location of the legend

#To display the histogram 
plt.show()


#1.2. Alternatively, comparison could be presented better using a histogram with 'non-overlapping' bars 
#plotting a non-overlapping histogram
plt.hist([Pitts_rating, LA_rating], 
          label=['Pittsburgh', 'Las Vegas'],
          color=['#9cd095', '#95c6d0'],
          alpha=0.8,
          linewidth=0.5, edgecolor='k',
          bins='auto')

#Adding a title 
plt.title('Distribution of Rating Scores for Las Vegas and Pittsburgh Businesses')
#Adding labels to histogram axes
plt.xlabel('Rating Score')
plt.ylabel('Frequency of Rating Score')
#Adding a legend to describe the histogram better
plt.legend(title='Business location:', loc='best')

#To display the histogram 
plt.show()



#2. Is the popularity of a given business (measured by reviews count) a good indication 
# of the qualify of its services (measured by star rating)?
#Create a scatter plot to assess the relationship between popularity and service quality 
# along 3 different classes of businesses: 'Health & Medical', 'Beauty & Spas', and 'Fashion'

#filtering data for the three types of businesses 
Health_filter1 = df['business type'] == 'Health & Medical'
Health_filter2 = df['service type'] == 'Health & Medical'
df_Health = df[Health_filter1 | Health_filter2]

Beauty_filter1 = df['business type'] == 'Beauty & Spas'
Beauty_filter2 = df['service type'] == 'Beauty & Spas'
df_Beauty = df[Beauty_filter1 | Beauty_filter2]

Fashion_filter1 = df['business type'] == 'Fashion'
Fashion_filter2 = df['service type'] == 'Fashion'
df_Fashion = df[Fashion_filter1 | Fashion_filter2]

#creating a scatterplot for each class of business to compare popularity to service quality for each
#plotting the data for health and medical industries
plt.scatter(df_Health['review_count'],           #specifying the data points to plot along the x-axis
            df_Health['stars'],                 #specifying the data points to plot along the y-axis
            label='Health & Medical',           #labeling the data points
            marker='o',                         #setting the marker type (circle-shaped)
            s=80,                              #setting the marker size 
            c='#90b0d0',                      #setting the marker color
            alpha=1)                          #setting the degree of transparency of marker 

#plotting the data for the beauty industries
plt.scatter(df_Beauty['review_count'],           
            df_Beauty['stars'],                
            label='Beauty & Spas',                    
            marker='o',                      
            s=80,                           
            c='#93d090',
            alpha=1)                       


#plotting the data for the fashion industries 
plt.scatter(df_Fashion['review_count'],
            df_Fashion['stars'],
            label='Fashion',
            marker='o',
            s=80,
            c='#f89d68',
            alpha=0.9)

#Adding a title to the scatter plot 
plt.title('The Relationship Between Popularity and Quality of Service')
#Adding labels to the scatterplot axes 
plt.xlabel('Popularity (measured by reviews frequency)')
plt.ylabel('Service Quality (measured by star rating)')
#Adding a legend
plt.legend(title='Business Type:', loc='best')

#To display the scatterplot 
plt.show()


#3. Create a pivot table that shows the mean star ratings for Health & Medical 
# businesses located in three different cities: 'Pittsburgh', 'Henderson', and 'Las Vegas'
#Plot a bar chart to compare the mean star ratings in the three cities 

#filtering data for Health & Medical 
Health_filter1 = df['business type'] == 'Health & Medical'
Health_filter2 = df['service type'] == 'Health & Medical'

#filtering data for cities 
Pitts_filter = df['city'] == 'Pittsburgh'
Henderson_filter = df['city'] == 'Henderson'
LA_filter = df['city'] == 'Las Vegas'

#creating a dataframe with the filtered data 
df_HealthByCity = df[(Health_filter1 | Health_filter2) & (Pitts_filter | Henderson_filter | LA_filter)]

#Creating a pivot table to calculate mean star rating per city 
pivot_HealthByCity = pd.pivot_table(df_HealthByCity, 
                                index=['city'], 
                                values=['stars'], 
                                aggfunc=np.mean)

#Plotting a bar chart to compare the means between cities
pivot_HealthByCity.plot(kind='bar',
                       color='#3235b3cf',
                       title='Average Star Rating for Health & Medical Businesses Per City',
                       legend=False,
                       width=0.35,           #setting the bars widths
                       figsize=(7,6),        #width x height
                       fontsize=8)         

#labeling bar chart axes 
plt.xlabel('Cities', fontsize=12)
plt.ylabel('Average Star Rating', fontsize=12)
#adjusting the tick label's rotation
plt.xticks(rotation=25)
#adjusting the y-axis scaler 
plt.ylim(0, 5)

#To display the bar chart 
plt.show()



#Part Six: Writing and/or Updating Excel Files 

#1. Writing a New Excel File With a Filtered Dataframe
#Say we want to filter data for businesses classified as either restaurants or bars,
# store them into a dataframe, and write a new Excel file with that filtered dataframe

#First, filtering data for 'Restaurants' or 'Bars' 
business_filter = df['business type'].isin(['Restaurants', 'Bars']) 
service_filter = df['service type'].isin(['Restaurants', 'Bars'])
df_RnB = df[business_filter | service_filter]

#Writing the dataframe, df_RnB, into the file 'yelp_filtered.xlsx' 
df_RnB.to_excel('yelp_filtered.xlsx',                #specifying file path and/or name
                sheet_name='Restaurants and Bars',   #specifying sheet name
                index=False)                         #removing the index coloumn

#in order to preview the file created 
try:
    df_newfile = pd.read_excel('yelp_filtered.xlsx')
    #to preview the first 5 enteries in the new file 
    print(df_newfile.head(), end='\n\n')
except:
    print('Error: file does not exist or cannot be accessed.')



#2. Writing New Data into an Existing Excel File

#This time I'll extract restaurants and bars specifically located in Las Vegas, store it into
#a new dataframe, df_LA_RnB, and write it into a new sheet in the 'yelp_filtered.xlsx' file

#filtering data for LA restaurants and bars 
LA_filter = df_RnB['city'] == 'Las Vegas'
df_LA_RnB = df_RnB[LA_filter]

#Writing dataframe df_LA_RnB into a new sheet 
with pd.ExcelWriter('yelp_filtered.xlsx', engine='openpyxl', mode='a') as writer: 
    df_LA_RnB.to_excel(writer, sheet_name='LA Restaurants and Bars', index=False)

#to preview the data in the new sheet
try: 
    xl = pd.ExcelFile('yelp_filtered.xlsx')
    df_newsheet = xl.parse('LA Restaurants and Bars')
    #to preview the first 5 enteries 
    print(df_newsheet.head(), end='\n\n')
except:
    print('Error: sheet does not exist or cannot be accessed.')



#3. Appending Data into Existing Sheets in an Existing Excel File

#Lastly, I'll filter the data further, say to extract businesses classified as 
#'Hotels & Travel' as well as those in LA only, and append the filtered data  
#into the two sheets I created earlier in the 'yelp_filtered.xlsx' file

#filtering data for 'Hotels & Travel' 
Hotels_filter1 = df['business type'] == 'Hotels & Travel'
Hotels_filter2 = df['service type'] == 'Hotels & Travel'
df_Hotels = df[Hotels_filter1 | Hotels_filter2]

#extracting those in LA 
LA_filter = df_Hotels['city'] == 'Las Vegas'
df_LA_Hotels = df_Hotels[LA_filter]

#Appending the two dataframes, df_Hotels and df_LA_Hotels, into the two existing sheets
with pd.ExcelWriter('yelp_filtered.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer: 
    #specifying the workbook with the existing data
    writer.book = openpyxl.load_workbook('yelp_filtered.xlsx')
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    #appending the dataframe df_Hotels to the 'Restaurants and Bars' sheet
    df_Hotels.to_excel(writer, 
                    sheet_name='Restaurants and Bars',        #specifying the sheet to append to 
                    index=False,                              #to remove unnecessary index coloumn 
                    header=False,                             #to remove unnecessary coloumn headers (of df_Hotels)
                    startrow=(len(df_RnB)+1))                 #specifying the starting row for appending new data


    #appending the dataframe df_LA_Hotels to the 'LA Restaurants and Bars' sheet
    df_LA_Hotels.to_excel(writer,
                        sheet_name='LA Restaurants and Bars',
                        index=False,
                        header=False,
                        startrow=(len(df_LA_RnB)+1))


#Renaming the sheets to match current data
#first, loading the file into a workbook object
workbook_obj = openpyxl.load_workbook('yelp_filtered.xlsx')
#Getting the original sheet names
sheet_1 = workbook_obj['Restaurants and Bars']
sheet_2 = workbook_obj['LA Restaurants and Bars']

#Changing the sheet names 
sheet_1.title = 'Leisure Businesses'
sheet_2.title = 'LA Leisure Businesses'

#saving the workbook with the new updates
workbook_obj.save('yelp_filtered.xlsx')


#To check if the data were appended successfully (& the other updates)
xl = pd.ExcelFile('yelp_filtered.xlsx')
df_sheet1 = xl.parse('Leisure Businesses')
df_sheet2 = xl.parse('LA Leisure Businesses')

#displaying the last 5 enteries in the first sheet
print('Last 5 enteries in sheet \'Leisure Businesses\':')
print(df_sheet1.tail(), end='\n\n')

#displaying the last 5 enteries in the second sheet
print('Last 5 enteries in sheet \'LA Leisure Businesses\':')
print(df_sheet2.tail())

#END