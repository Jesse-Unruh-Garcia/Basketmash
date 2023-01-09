# This is a sample Python script.
import openpyxl
import math
# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.


def main():
    # Use a breakpoint in the code line below to debug your script.
    firstPName = "DeMar DeRozan"
    p1Row = -1
    secondPName = "Luka Doncic"
    p2Row = -1
    middlePlayer = []

    workBook = openpyxl.load_workbook("NBAStats2022-2023.xlsx")
    workSheet = workBook["Sheet1"]

    # Start by fetching 1st/2nd Player Stats
    for collumn in workSheet['A']:
        if collumn.value == firstPName:
            p1Row = collumn.row
        if collumn.value == secondPName:
            p2Row = collumn.row
    for i in range(7):
        if i>0:
            p1Stat = workSheet[p1Row][i].value
            p2Stat = workSheet[p2Row][i].value
            middle = (p1Stat+p2Stat)/2
            middlePlayer.append(middle)
            #print(middle)

    #Create Mean of both stats

    euclidianDistanceScore = [0]*495
    stat = 0
    for col in workSheet.iter_cols(min_col=2,min_row=2,max_col=7,max_row=496):
        i=0
        for cell in col:
            check = cell.value
            added = (check - middlePlayer[stat])*(check-middlePlayer[stat])
            euclidianDistanceScore[i] += added
            i+=1
        stat+=1

    min = 10000000
    minInd = 0
    for i in range(len(euclidianDistanceScore)):
        euclidianDistanceScore[i] = math.sqrt(euclidianDistanceScore[i])
        if euclidianDistanceScore[i]<min:
            min = euclidianDistanceScore[i]
            minInd = i

    rowfin = workSheet[minInd+2]
    for a in range(len(euclidianDistanceScore)):
        row = workSheet[a+2]
        output = (str)(row[0].value) + ' has a score of '  + (str)(euclidianDistanceScore[a])
        print(output)
    print('And Finally, Our Winner')
    print(euclidianDistanceScore[minInd])
    print(minInd)
    for i in rowfin:
        print(i.value)





# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
