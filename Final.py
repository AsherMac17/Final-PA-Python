import csv
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from datetime import datetime

#Asks the user for 5 numbers and adds them together
def askUser():
    total = 0

    #This for loop runs 5 times and adds each number entered
    for i in range(5):
        number = int(input("Please enter a number: "))
        total = total + number

    print("The sum for the 5 numbers entered is:", total)


#Ask for 5 names and incomes and add them to final.csv
def askIncome():

    #This for loop runs 5 times and asks for a name and income, it is then added to a new line in final.csv
    for i in range(5):
        name = input("Please enter a name: ")
        income = input("Please enter their income: ")

        with open(r"C:\FinalExam\final.csv", "a")as file:
            file.write(name + "," + income + "\n")


#Create an Excel file with a pie chart
def excelPie():

    data = []

    with open(r"C:\FinalExam\final.csv", "r")as file:
        reader = csv.reader(file)

        for row in reader:
            data.append([row[0], int(row[1])])

    workbook = Workbook()
    sheet = workbook.active

    for row in data:
        sheet.append(row)

    today = datetime.now().strftime("%B %d, %Y")
    today = today.replace(" 0", " ")

    #Creates the pie chart
    chart = PieChart()

    #Selects the income numbers from column B for the pie chart
    chart_data = Reference(sheet, min_col=2, min_row=1, max_row=9)

    #Selects the names from column A to use as pie chart labels
    chart_labels = Reference(sheet, min_col=1, min_row=1, max_row=9)

    #Adds the income data to the pie chart
    chart.add_data(chart_data, from_rows=False, titles_from_data=False)

    #Adds the names as labels for the pie chart
    chart.set_categories(chart_labels)

    #Sets the title of the pie chart to studentid and today's date
    chart.title = "ashmac9358 " + today

    #Adds the pie chart to the worksheet
    sheet.add_chart(chart, "A11")

    #Saves the Excel file as final.xlsx
    workbook.save(r"C:\FinalExam\final.xlsx")

#Create a vertical bar graph
def verticalBar():

    names = []
    income = []

    with open(r"C:\FinalExam\final.csv", "r")as file:
        reader = csv.reader(file)

        for row in reader:
            names.append(row[0])
            income.append(int(row[1]))

    today = datetime.now().strftime("B %d, %Y")
    today = today.replace(" 0", " ")

    plt.bar(names, income)

    plt.xlabel("Name")
    plt.ylabel("Income")
    plt.title("ashmac9358 " + today)

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.show()

#Functions
askUser()
askIncome()
excelPie()
verticalBar()
