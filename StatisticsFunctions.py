#-*- coding: utf-8 -*-

from xlrd import open_workbook
import xlsxwriter as xls
from collections import OrderedDict, Counter
from scipy import stats
from prettytable import PrettyTable as PT
import unicodedata

def dataRead(file):
    '''This function reads an xls file and creates a dictionary containing the variable names and the
    data stored in each one.
    INPUT: Xls file route (string).
    OUTPUT: Excel data stored in a dictionary (dict).'''
    book = open_workbook(file)
    sheet = book.sheet_by_index(0)
    cols = sheet.ncols
    data = OrderedDict()
    headers = sheet.row(0)
    for h in range(len(headers)):
        if isinstance(headers[h].value, unicode):
            data[unicodedata.normalize('NFKD', headers[h].value).encode('ascii','ignore')] = []
    for column in range(cols):
        col = sheet.col(column)
        key = col[0].value
        if isinstance(key, unicode):
            key = unicodedata.normalize('NFKD', key).encode('ascii', 'ignore')
        for i in range(1,len(col)):
            value = col[i].value
            if isinstance(value, unicode):
                value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore')
            data[key].append(value)
    return data


def exportResult(table, path):
    '''This function exports a table obtained by an statistical test to an .xlsx file in the designed route.
    INPUT: Table to export (list) and the path of the file ending with file_name.xlsx (string).
    OUTPUT: Xlsx file saved in the designed route.'''
    workbook = xls.Workbook(path)
    worksheet = workbook.add_worksheet()
    row = 0
    for i in range(len(table)):
        data = table[i]
        for j in range(len(data)):
            worksheet.write(row, j, data[j])
        row = row + 1
    workbook.close()
    print "Data saved"


def analyzeBy(data, groupBy):
    '''This function sorts a data dictionary in different dictionaries, one for each category in the grouping
     variable.
     INPUT: data is the dictionary containing the data names and values (dict).  groupBy is the name of the
            grouping variable (string).
     OUTPUT: The output is a dictionary containing several dictionaries, one for each grouping category (dict).'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(groupBy, basestring):
            print('Error: groupBy must be a string with the name of the variable you would want to group by the data.')
        else:
            groupList = data[groupBy]
            del data[groupBy]
            cat = Counter(groupList)
            categories = cat.keys()
            sortedData = OrderedDict()
            for i in range(len(categories)):
                sortedData[categories[i]] = OrderedDict()
            for i in range(len(data.keys())):
                for j in range(len(sortedData.keys())):
                    sortedData[sortedData.keys()[j]][data.keys()[i]] = []
            for i in range(len(groupList)):
                for j in range(len(data.keys())):
                    sortedData[groupList[i]][data.keys()[j]].append(data[data.keys()[j]][i])
    return sortedData



def pairedTtest(data, printSig, *measures):
    '''This function computes the paired T-test for pairs of measures from data dictionary.
    INPUT: data is the dictionary containing the data names and values (dict).  printSig is
           a boolean variable, True: the function only prints the significative results, False:
           the function prints all the values (bool).  *measures contain all the pairs of
           variables to compare (strings).
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(printSig, bool):
            print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
                   'the function prints all the results.')
        else:
            if len(measures) % 2 == 0:
                results = OrderedDict()
                for i in range(0,len(measures), 2):
                    testName = measures[i] + '/' + measures[i + 1]
                    res = stats.ttest_rel(data[measures[i]], data[measures[i + 1]])
                    results[testName] = res
                table_matrix = [['', 'Test Statistic', 'p-Value']]
                if printSig:
                    m = results.keys()
                    for k in range(len(m)):
                        pVal = results[m[k]][1]
                        if pVal < 0.05:
                            table_matrix.append([m[k], results[m[k]][0], results[m[k]][1]])
                else:
                    m = results.keys()
                    for k in range(len(m)):
                        table_matrix.append([m[k], results[m[k]][0], results[m[k]][1]])
                table = PT(table_matrix[0])
                for row in range(1,len(table_matrix)):
                    table.add_row(table_matrix[row])
                print table
            else:
                print('Error: Measures must be paired two by two')
    return table_matrix

def indepTtest(data, printSig, groupBy, *measures):
    '''This function computes the independent T-test for measures grouped by groupBy from data dictionary.
    INPUT: data is the dictionary containing the data names and values (dict).  printSig is a boolean
           variable, True: the function only prints the significative results, False: the function
           prints all the values (bool).  groupBy is a list that contains 3 values, the first is the
           grouping variable, the second and the third are the groups to differentiate (list).  *measures
           contain all the pairs of variables to compare (strings).
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(printSig, bool):
            print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
                   'the function prints all the results.')
        else:
            if not isinstance(groupBy, list) and len(groupBy) == 3:
                print('Error: groupBy must be a list with three elements, the first one is the variable of grouping,'
                      ' the second and the third are the groups to compare.')
            else:
                indexG1 = []
                indexG2 = []
                results = OrderedDict()
                for i in range(len(data[groupBy[0]])):
                    if data[groupBy[0]][i] == groupBy[1]:
                        indexG1.append(i)
                    elif data[groupBy[0]][i] == groupBy[2]:
                        indexG2.append(i)
                for i in range(len(measures)):
                    m1 = []
                    m2 = []
                    for g1 in range(len(indexG1)):
                        m1.append(data[measures[i]][g1])
                    for g2 in range(len(indexG2)):
                        m2.append(data[measures[i]][g2])
                    levene = stats.levene(m1, m2)
                    if levene[1] > 0.05:
                        testName = measures[i] + ' (' + groupBy[1] + '/' + groupBy[2] + ')'
                        res = stats.ttest_ind(m1, m2, equal_var = True)
                        results[testName] = [levene, res]
                    elif levene[1] < 0.05:
                        testName = measures[i] + ' (' + groupBy[1] + '/' + groupBy[2] + ')'
                        res = stats.ttest_ind(m1, m2, equal_var=False)
                        results[testName] = [levene, res]
                table_matrix = [['', 'Levene Statistic', 'Levene p-Value','Test Statistic', 'p-Value']]
                if printSig:
                    m = results.keys()
                    for k in range(len(m)):
                        pVal = results[m[k]][1][1]
                        if pVal < 0.05:
                            table_matrix.append([m[k], results[m[k]][0][0], results[m[k]][0][1], results[m[k]][1][0],
                                                 results[m[k]][1][1]])
                else:
                    m = results.keys()
                    for k in range(len(m)):
                        table_matrix.append([m[k], results[m[k]][0][0], results[m[k]][0][1], results[m[k]][1][0],
                                                 results[m[k]][1][1]])
                table = PT(table_matrix[0])
                for row in range(1, len(table_matrix)):
                    table.add_row(table_matrix[row])
                print table
    return table_matrix
