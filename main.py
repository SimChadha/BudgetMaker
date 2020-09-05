import pygsheets
import sys
from pygsheets import HorizontalAlignment

# set up spreadsheet to use
account = pygsheets.authorize(service_file=r'C:\Users\simar\Desktop\BudgetSpreasheetMaker/client_secret.json')
sheet = account.open('PygSheets Tester')
page = sheet.worksheet("index", 0)

# create initial budget cells
budget = input("Enter your monthly budget: ")
budgetCell = page.cell("B1")
if budget[0] != "$":  # Adds $ if the user has not already
    budgetCell.value = "$" + budget
else:
    budgetCell.value = budget
budgetCell.text_format["bold"] = True
budgetTextCell = page.cell("A1")
budgetTextCell.value = "Budget: "
budgetTextCell.text_format["bold"] = True
budgetCell.update()
budgetTextCell.update()

# tuples for the colors of the entire row for price, date, and description
priceColor = (224 / 255, 102 / 255, 102 / 255)
dateColor = (109 / 255, 158 / 255, 235 / 255, 0)
descColor = (142 / 255, 124 / 255, 195 / 255)
colorsTuple = (dateColor, priceColor, descColor)  # to use in a loop later in the addNewData() function

# Creates Date header
tuple1 = (4,
          0)  # positional tuple to get to cell 4 rows down and 0 to the right. Entering in the tuple without a variable was not working
dateHeader = budgetCell.neighbour(tuple1)
dateHeader.value = "Date"
dateHeader.text_format["bold"] = True
dateHeader.text_format["underline"] = True
dateHeader.color = dateColor
dateHeader.update()

# Price Header
priceHeader = dateHeader.neighbour("right")
priceHeader.value = "Amount"
priceHeader.text_format["bold"] = True
priceHeader.text_format["underline"] = True
priceHeader.color = priceColor
priceHeader.update()

# Description Header
descriptionHeader = priceHeader.neighbour("right")
descriptionHeader.value = "Description"
descriptionHeader.text_format["bold"] = True
descriptionHeader.text_format["underline"] = True
descriptionHeader.color = descColor
descriptionHeader.update()

# Total expenses header
tuple1 = (3, 3)
totalHeader = descriptionHeader.neighbour(tuple1)
totalHeader.value = "Total Expenses:"
totalHeader.text_format["bold"] = True
totalHeader.update()

# Actual Expenses Cell
totalCell = totalHeader.neighbour("right")
# totalCell.value = "=CONCAT(\"$\",SUM())"  # Makes Cell that will have sum of amounts
totalCell.value = ""

# Budget Outlook Cell (over or under budget)
tuple1 = (1, 0)
outlookCell = totalCell.neighbour(tuple1)
outlookCell.text_format["bold"] = True
outlookCell.update()

listOfCosts = []  # Array that will store all the costs user enters
currentSum = 0  # Will store the current *sum* of those costs

catchingUp = True  # Tells findNextRow if it should be adding up previous values or not. Should on;ly do so on first run through

# Function that finds the first available row that is not populated with budget data
def findNextRow():
    tempTuple = (1, 0)  # again, simply using a tuple without first making a variable doesn't work in the next line for some reason
    indexCell = dateHeader.neighbour(tempTuple)  # to start checking at the column directly below Date Header
    global catchingUp
    while True:
        global currentSum
        if (indexCell.value == ""):
            catchingUp = False  # Done adding in old numbers once we start adding new ones
            return indexCell
        else:
            if catchingUp:
                currentSum = currentSum + int(indexCell.neighbour("right").value[1:])  # Makes sure data from previous iterations are uses in current sum
                print("INSIDE CATCHING: " + str(currentSum))
            indexCell = indexCell.neighbour(tempTuple)  # moves index to row directly below last checked one


# Will insert the addition into the text before the first ")" in text
def insetBeforeP(origText, addition):
    withoutP = origText[0:origText.index(")")]
    withoutP = withoutP + addition + "))"
    return withoutP


# def isFirstSumAddition():
#     if len(totalCell.value) == 18:
#         return True
#     return False

# Sums all values given a list of strings that can be parsed to int
def sumStringList(list):
    sum = 0
    for i in range(len(list)):
        sum = sum + int(list[i])
    return sum

firstTimeAddingSum = True
# Takes in User data to add to sheet
def addNewData(newLine):
    while True:
        global firstTimeAddingSum
        # try:
        if newLine == "exit" or newLine == "Exit":
            sys.exit()
        listOfAdditions = newLine.split(",")
        blankCell = findNextRow()
        for i in range(3):
            global currentSum
            blankCell.value = listOfAdditions[i]
            blankCell.color = colorsTuple[i]
            if i == 1:  # Special actions need to take place when adding the *price* user data
                if listOfAdditions[i][0] == '$':  # If the user accidentally adds in the "$",
                    listOfAdditions[i] = listOfAdditions[i][1:]  # Remove the '$'
                blankCell.value = "$" + listOfAdditions[i]  # Add '$' in
                listOfCosts.append(listOfAdditions[i])
                currentSum = currentSum + int(listOfCosts[-1])
                totalCell.value = "$" + str(currentSum)
            blankCell = blankCell.neighbour("right")
        outlookCell.value = int(
            budgetCell.value[1:]) - currentSum  # Gets the value and removes the $ sign from totalCell
        if int(outlookCell.value) < 0:
            outlookCell.color = (204 / 255, 0, 0)  # Dark red
        if int(outlookCell.value) > 0:
            outlookCell.color = (106 / 255, 168 / 255, 79 / 255)  # Dark Green
        if int(outlookCell.value) == 0:
            outlookCell.color = (204 / 255, 65 / 255, 37 / 255)  # Lighter red
        break
    # except SystemExit:  # Allows exiting of program when needed since just one sys.exit() does not work
    # sys.exit()
    # except:
    # print("The format entered was not correct.")


timeToExit = False
while not (timeToExit):  # Main loop that allows users to continue to enter data and exit when needed
    print(currentSum)
    addNewData(input(
        "Enter your purchase in the following format: Date,Price(no$sign),Description of Purchase\nEnter \"Exit\" to stop adding\n"))

# page.update_value("A1","")
