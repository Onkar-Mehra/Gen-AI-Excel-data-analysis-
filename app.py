import re
import numpy as np
import os
import openai
from flask import Flask,render_template,url_for,request, redirect,jsonify
from flask_cors import CORS
from flask import session
from flask_sqlalchemy import SQLAlchemy
from datetime import date
from requests.models import PreparedRequest
import os
import pandas as pd

app = Flask(__name__)
CORS(app)

# App secret key
app.secret_key = os.getenv("APP_SECRET_KEY") 

filename1='filename.xlsx'
xls=pd.ExcelFile(filename1)
sheetnames=xls.sheet_names
number_of_sheets = len(sheetnames)
separator=', '
excelsheetnames=separator.join(sheetnames)
print(f"Sheet Names: {sheetnames}")
print(f"Number of Sheets: {number_of_sheets}")
excelsheetnames
opt=input("Press 1 for your query to be addressed in a graphical format, and press 2 for it to be addressed in a textual format-")
print('\n Reading and processing the data.... \n Wait for sometime before entering the query....')
file=pd.read_excel(filename1,excelsheetnames)
columns=file.columns.tolist()
first_5_rows = file.head(5).to_string(index=False)
text=f"Column Names: {', '.join(columns)}\nFirst 5 Rows:\n{first_5_rows}"
text

if opt=='1':
    def read_excel_and_concat_unique_values(file_path):
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)

        # Initialize an empty dictionary to store the result
        result_dict = {}

        # Iterate through each column in the DataFrame
        for column in df.columns:
            # Check if the number of unique values in the column is less than 30
            if df[column].nunique() < 30:
                # Store the column name and its unique values in the dictionary
                result_dict[column] = df[column].unique()

        # Format the result as a string
        result_string = "\n".join([f"{column}: {values}" for column, values in result_dict.items()])

        return result_string

    result = read_excel_and_concat_unique_values(filename1)

    input_query_message="You are a system  that reads the prompt/query and generate graph based on the excel file name %s and sheet name %s."%(filename1,sheetnames)+" You have to understand and read the query given by user then start identifying the process of suitable columns to resolve the query of user.  You have been provided the column headers of the data and the data of first five rows according to which you have to choose the most appropriate column headers as per the best assupmtion. Please comprehend the interconnection among these columns and identify related pairs that can aid in resolving the query. Choose appropriate columns for the x-axis and y-axis to enable graph plotting based on their correlation. These are the column headings and the data for first 5 rows- "+text+" After this You are  provided with a list of columns from the Excel sheet, each containing unique values less than or equal to 30. You have to carefully and properly understand and analyse and compare the attributes mentioned in the query with the provided column names and unique values. If a match is found, choose only those columns that where the similar values are found for the development of the graph. The data for column names and its unique values are- "+result +". Compulsorily traverse through the column names and the unique values present in them. Choose the best and appropriate columns that later on will be used to plot the graph according to the query and resolve the query and print() the column names at the end. Print only unique values without any introduction and explanation. You have to "+"Find the most appropraite columns for the creation of graph depending on the prompt, column headers and data of first five rows and the column headers with the unique values that will resolve the user's query efficiently and plot the graph. In case it's challenging to identify precise columns to resolve the query, then again properly read and understand all the column headers and the data of first five rows-"+text+" and the list of columns with the unique values belonging to the same data- "+result+"to select the columns that seem most likely to address the query or determine a solution that involves column names provided, aiming to best approximate a resolution. Avoid sending a response indicating no columns found and refrain from making assumptions about columns not provided in the given column names and its data; limit selections to the given columns and their data."
    original_query= input('Enter the prompt (The prompt should have atleast 4 characters)- ')
    if len(original_query)>=4:
        original_message= original_query 
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                    engine= "*********",
                                                    model="*********",
                                                    messages=[{"role": "system", "content": input_query_message},
                                                              {"role": "user", "content": original_message }],
                                                    temperature = 0,
                                                    top_p = 0.95,
                                                    frequency_penalty = 0,
                                                    presence_penalty = 0,
                                                    stop = None
                                                    )
        oldprompt = response['choices'][0]['message']['content']
        print("Prompt generated as per your query :",oldprompt)

    mod_message='You are a system that accepts the response on the question that whether you are satified with the chosen column names that will be later on used to develop the graph and creates a new prompt using original prompt %s and specified column names %s .'%(original_query,oldprompt)+' If any additions or deletions has to be done in the column name, the user will  tell u what has to be deleted or added. If the user is satisfied with the present columns mentioned then proceed with these column names otherwise remove or add the column names as per the response of users. If any new column has to be added, to make sure that the correct column name is added, you are provided with the column header and the data of first five rows'+text+' Then at the end using original prompt and the new or old column names depending on the user, create a new efficient and profficient prompt for the creation of graph making sure that there are attributes to be presented at x-axis and and attributes to be mentioned at the y-axis'
    mod_input= input('Are you satisfied with created prompt? If no, then enter the modifications to be done- ')
    if len(mod_input)>=3:
        system_message= mod_input  
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                    engine= "*********",
                                                    model="*********",
                                                    messages=[{"role": "system", "content": mod_message},
                                                              {"role": "user", "content": mod_input }],
                                                    temperature = 0,
                                                    top_p = 0.95,
                                                    frequency_penalty = 0,
                                                    presence_penalty = 0,
                                                    stop = None
                                                    )
        generated_text = response['choices'][0]['message']['content']
        print("New prompt by AI :",generated_text)
    final_message="You are a system that reads the data of excel file named %s and sheet name %s and accepts the query from the user to generate graph or chart based on the uploaded CSV/Excel file. Based on the query given by user, you generate only the python code without any kind of introduction, appology or instruction and generate a graph that can be a bar chart, line graph chart, pie chart or column chart with proper labelling and legends that will be the most efficient in making the user understand."%(filename1,excelsheetnames)+"This is the data for you to read and then generate the code accordingly using desired column names as per your understanding-"+text+"You use the columns names and the data of first five rows to understand that what the excel sheet is all about and then start developing the code using the most appropriate columns and their data in order to show the best graph depending on the query"+"For plotting the graph strictly adhere to the given rules- 1) you generate only the python code without any kind of introduction, apology or instruction. \n 2)generate a graph that can be a bar chart, line graph chart, pie chart or column chart with proper labelling and legends that will be the most efficient in making the user understand. \n 3) If the final output is in the form of dataframe then print it at the end by importing library Ipython.display otherwise simply plot the graph. 4) To avoid 'KeyError', take the alternative steps. \n 5) Make the chart highly detailed and aesthetic. \n 6) (Important) At the end, read and analyse the plotted chart and print() the key features in textual form that have been observed from the chart using print(). \n 7) Make all the graphs with many labels on x-axis big enough so that the labels do not overlap each other and are clearly visible. The labels on x-axis are clearly visible. 8) When you have a large number of categories on the x-axis of a graph and they overlap, you can address this issue by either increasing the spacing between categories to make them distinct, or by grouping and representing them based on their count. To do this, you can calculate bins or categories based on the total number of categories, making the graph more readable and avoiding overlap. When faced with numerous x-axis categories, you can convert them into numerical representations and display them using either their corresponding numbers, counts, frequencies, or organize them into bins for a more manageable presentation. \n 8) Show labels on the x-axis in vertical/slant direction. \n 10) Keep figsize=(16,8) \n 11) If the query asks for qarterly based data, then convert the date into quarter. \n 12) If the query asks for monthly based data, then convert the date into months. Show the date into textual format on x-axis. Show the data in the form of months in textual format. For example if there is 01, then change it into Jan. \n 13) If the query asks for yearly based data, then convert data into years. \n 14) Wherever necessary keep labelsize=12 \n 15) Move chart legends to the right side and do not keep them in between. \n 16) (IMPORTANT) Ensure that 'TypeError: Invalid object type at position 0' does not occur therefore read the excel accordingly. \n 17) If the query asks for monthly data, then show the dates on x-axis in words. \n 18) If the query requests yearly data, verify whether there is a sufficient amount of data available to present it in yearly terms. If not, display the data in monthly format using words. Never show years in decimals rather show the data in the form of quarters. \n 19) (Highly Important) Add data labels in the chart to let the user know that at which point where the data stands. Generate code accordingly so that the data labels are visble on the graph. \n 20) (IMPORTANT) Do not include any instruction, appology or introduction and extract and print only the code that will be run. \n 21) (HIGHLY IMPORtant) Add the code compulsorily to add annotations in the graph. Compulsorily add annotations in all the types of graph to make it more understanding. Create a code to create well annotated graph. \n 22) (Highly important) Generate code that is intended for execution using the `exec` function. The code should include print statements for any warnings and key features, ensuring that the `exec` function will directly print these messages. Please refrain from writing the messages directly; instead, incorporate them into print statements to facilitate direct printing when the code is executed with `exec`.\n 23) (Highly Important) When writing code, avoid assuming the presence of any columns in the excel sheet. If you want to create a new column, use the columns that are already present in the excel sheet. ADD ANNOTATIONS TO THE GRAPH. \n 24) Avoid the error 'division by zero' and choose an alternative method. Do not let the error 'Division by zero' occur. \n 25) Avoid invalid object type at position 0. \n 26) Avoid the error 'ValueError: Image size of 230555x500 pixels is too large. It must be less than 2^16 in each direction.'. Set the image size accordingly so that the graph image can be printed easily without any error. \n 27) (Highly Important) Choose the graph type that is best suited for representing the data. Data can have multiple attributes to be represented on a single axis. Therefore choose the most suitable chart that can be bar chart, line chart, column chart, pie chart for plotting the data. The graph can be double, triple or depending on the number of attributes, the graph can be plotted. If there are more than one attribute to be plotted on a single axis then write the code in such a way that the multiple charts can be plotted.  \n 28) (Highly Important) Avoid this error - Cannot subset columns with a tuple with more than one element. Use a list instead. \n 29) (IMPORTANT) Make sure that the code generated can be easily executed uing exec() function. \n 30) while creating the code, make sure that all the varaiables such as barWidth are clearly defined so that it does not throw any error. \n 31) (HIGHLY IMPORTANT & HIGH PRIORITY) Compulsorily add annotations to the graph."

    if len(generated_text)>=4:
        system_message= generated_text 
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                engine= "*********",
                                                model="*********",
                                                messages=[{"role": "system", "content": final_message},
                                                          {"role": "user", "content": system_message }],
                                                temperature = 0,
                                                top_p = 0.95,
                                                frequency_penalty = 0,
                                                presence_penalty = 0,
                                                stop = None
                                                )
        code = response['choices'][0]['message']['content']
        print("Code response by AI :",code)
        pattern = r'```python\n(.*?)\n```'
        match = re.search(pattern, code, re.DOTALL)
        if match:
            extracted_code = match.group(1)
            print(extracted_code)
            exec(extracted_code)
        else:
            print("Error occurred")
            
    relevant_query_message='You are an intelligent system that reads the prompt entered by the user that is inteneded to ask questions related to the table in the excel file and you generate 5 questions related or based on the prompt and based on the column headers and the data of first 5 rows. These are the column headers and the data of first 5 rows'+text+'The 5 prompts/questions that you generate can be similiar/related to the entered prompt and can be based on the column headers and the data of first 5 rows given earlier.'
    if len(original_query)>=4:
        original_message= original_query 
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                engine= "*********",
                                                model="*********",
                                                messages=[{"role": "system", "content": relevant_query_message},
                                                          {"role": "user", "content": original_message }],
                                                temperature = 0,
                                                top_p = 0.95,
                                                frequency_penalty = 0,
                                                presence_penalty = 0,
                                                stop = None
                                                )
        relevant_questions = response['choices'][0]['message']['content']
        print("Relevant questions as per your query-\n",relevant_questions)
            
elif opt=='2':
    text_query_message="You are a system that reads the data of excel file named %s and sheet name %s and accepts the query from the user to generate or print() a textual answer based on the uploaded CSV/Excel file. Based on the query given by user, you generate only the python code without any kind of introduction, appology or instruction and generate an answer in textual format that will be the most efficient in making the user understand."%(filename1,excelsheetnames)+"This is the data for you to read and then generate the code accordingly using desired column names as per your understanding-"+text+"You use the columns names and the data of first five rows to understand that what the excel sheet is all about and then start developing the code using the most appropriate columns and their data in order to show the answer in textual format. For printing the answer in textual format strictly adhere to the given rules- 1) you generate only the python code without any kind of introduction, apology or instruction. \n 2) Print() the answer in textual format that will be the most efficient in making the user understand. \n 3) To avoid 'KeyError', take the alternative steps. \n 4) Make the answer highly detailed. \n 5) If the query asks for qarterly based data, then convert the date into quarter. \n 6) If the query asks for monthly based data, then convert the date into months. Show the date into textual format. Show the data in the form of months in textual format. For example if there is 01, then change it into Jan. \n 7) If the query asks for yearly based data, then convert data into years. \n 8) (IMPORTANT) Ensure that 'TypeError: Invalid object type at position 0' does not occur therefore read the excel accordingly. \n 9) If the query asks for monthly data, then show the dates in words. \n 10) If the query requests yearly data, verify whether there is a sufficient amount of data available to present it in yearly terms. If not, display the data in monthly format using words. Never show years in decimals rather show the data in the form of quarters. \n 11) (IMPORTANT) Do not include any instruction, appology or introduction and extract and print only the code that will be run. \n 12) (Highly important) Generate code that is intended for execution using the `exec` function. The code should include print statements for any warnings and key features, ensuring that the `exec` function will directly print these messages. Please refrain from writing the messages directly; instead, incorporate them into print statements to facilitate direct printing when the code is executed with `exec`. \n 13) (Highly Important) When writing code, avoid assuming the presence of any columns in the excel sheet. If you want to create a new column, use the columns that are already present in the excel sheet \n 14) (Important) Properly the numbers should be displayed and not in the form of exponential notation. Do not display numbers in the form of exponential notation. \n 15) At the end print() all the columns that were used to print the anser in textual format."
    text_query= input('Enter the prompt (The prompt should have atleast 4 characters)- ')
    if len(text_query)>=4:
        original_message= text_query 
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                engine= "*********",
                                                model="*********",
                                                messages=[{"role": "system", "content": text_query_message},
                                                          {"role": "user", "content": original_message }],
                                                temperature = 0,
                                                top_p = 0.95,
                                                frequency_penalty = 0,
                                                presence_penalty = 0,
                                                stop = None
                                                )
        textual_answer = response['choices'][0]['message']['content']
        print("Textual answer as per your query :",textual_answer)
        pattern = r'```python\n(.*?)\n```'
        match = re.search(pattern, textual_answer, re.DOTALL)
        if match:
            extracted_code = match.group(1)
            print(extracted_code)
            exec(extracted_code)
        else:
            print("Error occurred")
            
    relevant_query_message='You are an intelligent system that reads the prompt entered by the user that is inteneded to ask questions related to the table in the excel file and you generate 5 questions related or based on the prompt and based on the column headers and the data of first 5 rows. These are the column headers and the data of first 5 rows'+text+'The 5 prompts/questions that you generate can be similiar/related to the entered prompt and can be based on the column headers and the data of first 5 rows given earlier.'
    if len(text_query)>=4:
        original_message= text_query 
        openai.api_type = 'azure'
        openai.api_base ="*********"
        openai.api_version = '*********'
        openai.api_key = '*********'
        response = openai.ChatCompletion.create(
                                                engine= "*********",
                                                model="*********",
                                                messages=[{"role": "system", "content": relevant_query_message},
                                                          {"role": "user", "content": original_message }],
                                                temperature = 0,
                                                top_p = 0.95,
                                                frequency_penalty = 0,
                                                presence_penalty = 0,
                                                stop = None
                                                )
        relevant_questions = response['choices'][0]['message']['content']
        print("Relevant questions as per your query-\n",relevant_questions)
