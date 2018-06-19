#-*- coding: utf-8 -*-

from xlrd import open_workbook
from collections import OrderedDict, Counter, namedtuple
from scipy import stats
from prettytable import PrettyTable as PT
import unicodedata
import xlsxwriter as xls
import itertools
import pyvttbl as pt

# -----DATA IMPORT AND EXPORT FUNCTIONS---------------------------------------------------------------------------------


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

# -----T-TEST FUNCTIONS-------------------------------------------------------------------------------------------------


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
                table_matrix = [['Paired T-test', 'Test Statistic', 'p-Value']]
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
                table_matrix = [['Independent T-test', 'Levene Statistic', 'Levene p-Value','Test Statistic',
                                 'p-Value']]
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

# -----CORRELATION TEST FUNCTIONS---------------------------------------------------------------------------------------


def pearsonCorrel(data, printSig, *measures):
    '''This function computes the Pearson correlation over all the possible pairs of the variables included.
    INPUT: data is the dictionary containing the data names and values (dict).  printSig is a boolean
           variable, True: the function only prints the significative results, False: the function
           prints all the values (bool). *measures contain all the variables to compare (strings).
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(printSig, bool):
            print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
                   'the function prints all the results.')
        else:
            if not  len(measures) >= 2:
                print('Error: At least two measures are necessary to compute correlation.')
            else:
                pairs = list(itertools.combinations(measures, 2))
                results = OrderedDict()
                for i in range(len(pairs)):
                    testName = pairs[i][0] + '/' + pairs[i][1]
                    res = stats.pearsonr(data[pairs[i][0]], data[pairs[i][1]])
                    results[testName] = res
                table_matrix = [['Pearson correlation', 'Correl. coefficient', 'p-Value']]
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
    return table_matrix

# -----OTHER TEST FUNCTIONS---------------------------------------------------------------------------------------------


def normalityTest(data, printSig, *measures):
    '''This function computes the normality test for the variables included.
    INPUT: data is the dictionary containing the data names and values (dict).  printSig is
           a boolean variable, True: the function only prints the significative results, False:
           the function prints all the values (bool).  *measures contain all the variables to
           compute the test over (strings).
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(printSig, bool):
            print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
                   'the function prints all the results.')
        else:
            results = OrderedDict()
            for i in range(len(measures)):
                testName = measures[i]
                res = stats.normaltest(data[measures[i]])
                results[testName] = res
            table_matrix = [['Normality test', 'Test Statistic', 'p-Value']]
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
            for row in range(1, len(table_matrix)):
                table.add_row(table_matrix[row])
            print table
    return table_matrix

# -----GROUPED T-TEST FUNCTIONS-----------------------------------------------------------------------------------------


def analyzeBy(data, sortBy):
    '''This function sorts a data dictionary in different dictionaries, one for each category in the grouping
     variable.
     INPUT: data is the dictionary containing the data names and values (dict).  sortBy is the name of the
            grouping variable (string).
     OUTPUT: The output is a dictionary containing several dictionaries, one for each grouping category (dict).'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(sortBy, basestring):
            print('Error: sortBy must be a string with the name of the variable by wich you would want to group the'
                  ' data.')
        else:
            tempData = data.copy()
            groupList = tempData[sortBy]
            del tempData[sortBy]
            cat = Counter(groupList)
            categories = cat.keys()
            sortedData = OrderedDict()
            for i in range(len(categories)):
                sortedData[categories[i]] = OrderedDict()
            for i in range(len(tempData.keys())):
                for j in range(len(sortedData.keys())):
                    sortedData[sortedData.keys()[j]][tempData.keys()[i]] = []
            for i in range(len(groupList)):
                for j in range(len(tempData.keys())):
                    sortedData[groupList[i]][tempData.keys()[j]].append(tempData[tempData.keys()[j]][i])
    return sortedData


def groupedPairedTtest(data, sortBy, printSig, *measures):
    '''This function computes the paired T-test for pairs of measures from data dictionary.
        INPUT: data is the dictionary containing the data names and values (dict).  printSig is
               a boolean variable, True: the function only prints the significative results, False:
               the function prints all the values (bool).  *measures contain all the pairs of
               variables to compare (strings).
        OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    sortedData = analyzeBy(data, sortBy)
    if not isinstance(printSig, bool):
        print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
               'the function prints all the results.')
    else:
        if len(measures) % 2 == 0:
            fullResults = OrderedDict()
            for i in range(len(sortedData.keys())):
                groupName = sortedData.keys()[i]
                tempData = sortedData[sortedData.keys()[i]]
                results = OrderedDict()
                for j in range(0, len(measures), 2):
                    testName = measures[j] + '/' + measures[j + 1]
                    res = stats.ttest_rel(tempData[measures[j]], tempData[measures[j + 1]])
                    results[testName] = res
                fullResults[groupName] = results
            table_matrix = [['', 'Paired T-test', 'Test Statistic', 'p-Value']]
            for i in range(len(fullResults.keys())):
                if printSig:
                    results = fullResults[fullResults.keys()[i]]
                    m = results.keys()
                    for k in range(len(m)):
                        pVal = results[m[k]][1]
                        if pVal < 0.05:
                            if k == 0:
                                table_matrix.append([fullResults.keys()[i], m[k], results[m[k]][0], results[m[k]][1]])
                            else:
                                table_matrix.append(['', m[k], results[m[k]][0], results[m[k]][1]])
                else:
                    results = fullResults[fullResults.keys()[i]]
                    m = results.keys()
                    for k in range(len(m)):
                        if k == 0:
                            table_matrix.append([fullResults.keys()[i], m[k], results[m[k]][0], results[m[k]][1]])
                        else:
                            table_matrix.append(['', m[k], results[m[k]][0], results[m[k]][1]])
            table = PT(table_matrix[0])
            for row in range(1, len(table_matrix)):
                table.add_row(table_matrix[row])
            print table
        else:
            print('Error: Measures must be paired two by two')
    return table_matrix


def groupedIndepTtest(data, sortBy, printSig, groupBy, *measures):
    '''This function computes the paired T-test for pairs of measures from data dictionary.
        INPUT: data is the dictionary containing the data names and values (dict).  printSig is
               a boolean variable, True: the function only prints the significative results, False:
               the function prints all the values (bool).  *measures contain all the pairs of
               variables to compare (strings).
        OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    sortedData = analyzeBy(data, sortBy)
    if not isinstance(printSig, bool):
        print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
               'the function prints all the results.')
    else:
        if not isinstance(groupBy, list) and len(groupBy) == 3:
            print('Error: groupBy must be a list with three elements, the first one is the variable of grouping,'
                  ' the second and the third are the groups to compare.')
        else:
            fullResults = OrderedDict()
            for i in range(len(sortedData.keys())):
                groupName = sortedData.keys()[i]
                tempData = sortedData[sortedData.keys()[i]]
                indexG1 = []
                indexG2 = []
                results = OrderedDict()
                for j in range(len(tempData[groupBy[0]])):
                    if tempData[groupBy[0]][j] == groupBy[1]:
                        indexG1.append(j)
                    elif tempData[groupBy[0]][j] == groupBy[2]:
                        indexG2.append(j)
                for j in range(len(measures)):
                    m1 = []
                    m2 = []
                    for g1 in range(len(indexG1)):
                        m1.append(tempData[measures[j]][g1])
                    for g2 in range(len(indexG2)):
                        m2.append(tempData[measures[j]][g2])
                    levene = stats.levene(m1, m2)
                    if levene[1] > 0.05:
                        testName = measures[j] + ' (' + groupBy[1] + '/' + groupBy[2] + ')'
                        res = stats.ttest_ind(m1, m2, equal_var=True)
                        results[testName] = [levene, res]
                    elif levene[1] < 0.05:
                        testName = measures[j] + ' (' + groupBy[1] + '/' + groupBy[2] + ')'
                        res = stats.ttest_ind(m1, m2, equal_var=False)
                        results[testName] = [levene, res]
                    fullResults[groupName] = results
            table_matrix = [['', 'Independent T-test', 'Levene Statistic', 'Levene p-Value', 'Test Statistic',
                            'p-Value']]
            for i in range(len(fullResults.keys())):
                if printSig:
                    results = fullResults[fullResults.keys()[i]]
                    m = results.keys()
                    for k in range(len(m)):
                        pVal = results[m[k]][1][1]
                        if pVal < 0.05:
                            if k == 0:
                                table_matrix.append([fullResults.keys()[i], m[k], results[m[k]][0][0],
                                                     results[m[k]][0][1], results[m[k]][1][0], results[m[k]][1][1]])
                            else:
                                table_matrix.append(['', m[k], results[m[k]][0][0],
                                                     results[m[k]][0][1], results[m[k]][1][0], results[m[k]][1][1]])
                else:
                    results = fullResults[fullResults.keys()[i]]
                    m = results.keys()
                    for k in range(len(m)):
                        if k == 0:
                            table_matrix.append([fullResults.keys()[i], m[k], results[m[k]][0][0],
                                                 results[m[k]][0][1], results[m[k]][1][0], results[m[k]][1][1]])
                        else:
                            table_matrix.append(['', m[k], results[m[k]][0][0],
                                                 results[m[k]][0][1], results[m[k]][1][0], results[m[k]][1][1]])
                table = PT(table_matrix[0])
                for row in range(1, len(table_matrix)):
                    table.add_row(table_matrix[row])
            print table
    return table_matrix


# -----ANOVA FUNCTIONS--------------------------------------------------------------------------------------------------


def repeatedMeasuresAnova(data, subID, conditionName, *measures):
    '''This function computes a ANOVA for repeated measures over the variables defined along
    with the condition factor.
    INPUT: data is the dictionary containing the data names and values (dict). subID is the name
           of the variable that codes the identifier of the subjects(string). conditionName is the
           name of the condition over you want to compute the ANOVA (string). *measures contain
           pairs of variable / condition over you want to compute the ANOVA.
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(subID, basestring):
            print ('Error: subID must be a string containing the name of the variable with the subjects ID.')
        else:
            if not isinstance(conditionName, basestring):
                print ('Error: conditionName must be a string containing the name of the condition analyzed in the'
                       ' anova.')
            else:
                errorCount = 0
                for elem in range(len(measures)):
                    if not isinstance(measures[elem], tuple):
                        errorCount = errorCount + 1
                if errorCount != 0:
                    print('Error: measures must contain tuples with a data variable and an associated condition.')
                else:
                    Sub = namedtuple('Sub', ['Sub_id', 'measure', 'condition'])
                    df = pt.DataFrame()
                    for i in range(len(measures)):
                        meas = data[measures[i][0]]
                        for j in range(len(meas)):
                            df.insert(Sub(data[subID][j], meas[j], measures[i][1])._asdict())
                    aov = df.anova('measure', sub='Sub_id', wfactors=['condition'])
                    table_matrix = [['Source', '', 'Type III SS', 'df', 'SM', 'F', '.Sig'],

                                    [conditionName, 'Sphericity Assumed', aov[('condition',)]['ss'], aov[('condition',)]['df'],
                                     aov[('condition',)]['mss'], aov[('condition',)]['F'], aov[('condition',)]['p']],

                                    ['', 'Greenhouse-Geiser', aov[('condition',)]['ss'], aov[('condition',)]['df_gg'],
                                     aov[('condition',)]['mss_gg'], aov[('condition',)]['F_gg'], aov[('condition',)]['p_gg']],

                                    ['', 'Hyunh-Feldt', aov[('condition',)]['ss'], aov[('condition',)]['df_hf'],
                                     aov[('condition',)]['mss_hf'], aov[('condition',)]['F_hf'], aov[('condition',)]['p_hf']],

                                    ['', 'Box', aov[('condition',)]['ss'], aov[('condition',)]['df_lb'], aov[('condition',)]['mss_lb'],
                                     aov[('condition',)]['F_lb'], aov[('condition',)]['p_lb']],

                                    ['Error(' + conditionName + ')', 'Sphericity Assumed', aov[('condition',)]['sse'], aov[('condition',)]['dfe'],
                                     aov[('condition',)]['mse'], '-', '--'],

                                    ['', 'Greenhouse-Geiser', aov[('condition',)]['sse'], aov[('condition',)]['dfe_gg'],
                                     aov[('condition',)]['mse_gg'], '-', '--'],

                                    ['', 'Hyunh-Feldt', aov[('condition',)]['sse'], aov[('condition',)]['dfe_hf'],
                                     aov[('condition',)]['mse_hf'], '-', '--'],

                                    ['', 'Box', aov[('condition',)]['sse'], aov[('condition',)]['dfe_lb'],
                                     aov[('condition',)]['mse_lb'], '-', '--']]
                    table = PT(table_matrix[0])
                    for row in range(1, len(table_matrix)):
                        table.add_row(table_matrix[row])
                    print table
    return table_matrix


# -----POST HOC-TEST----------------------------------------------------------------------------------------------------


def repMeasBonferroniCorrect(data, printSig, *measures):
    '''This function computes the Bonferroni correction for pairwise  combination of measures
    from data dictionary.
    INPUT: data is the dictionary containing the data names and values (dict).  printSig is
           a boolean variable, True: the function only prints the significative results, False:
           the function prints all the values (bool).  *measures contain all the variables to
           compute the pairwise tests (strings).
    OUTPUT: The function prints a table in the terminal containing all the tests computed.'''
    if not isinstance(data, dict):
        print ('Error: data must be a dict. Use dataRead function to import your excel data.')
    else:
        if not isinstance(printSig, bool):
            print ('Error: printSig must be a bool. True: the function only prints the siginificative results/ False: '
                   'the function prints all the results.')
        else:
            n = len(measures)
            results = OrderedDict()
            for i in range(0, len(measures)):
                results[measures[i]] = []
                for j in range(0, len(measures)):
                    if j != i:
                        res = stats.ttest_rel(data[measures[i]], data[measures[j]])
                        results[measures[i]].append([measures[j], res])
            table_matrix = [['Bonferroni correction', 'Pairwise T-test', 'Test Statistic', 'p-Value']]
            if printSig:
                m = results.keys()
                for k in range(len(m)):
                    count = 0
                    for t in range(len(results[m[k]])):
                        pVal = results[m[k]][t][1][1]
                        if pVal < (0.05 / n):
                            count = count + 1
                            if count == 1:
                                table_matrix.append([m[k], results[m[k]][t][0], results[m[k]][t][1][0], results[m[k]][t][1][1]])
                            else:
                                table_matrix.append(['', results[m[k]][t][0], results[m[k]][t][1][0], results[m[k]][t][1][1]])
                table_matrix.append(['Sig if p-Value < ' + str(0.05 / n), '-', '--', '---'])
            else:
                m = results.keys()
                for k in range(len(m)):
                    for t in range(len(results[m[k]])):
                        if t == 0:
                            table_matrix.append(
                                [m[k], results[m[k]][t][0], results[m[k]][t][1][0], results[m[k]][t][1][1]])
                        else:
                            table_matrix.append(
                                ['', results[m[k]][t][0], results[m[k]][t][1][0], results[m[k]][t][1][1]])
                table_matrix.append(['Sig if p-Value < ' + str(0.05 / n), '-', '--', '---'])
            table = PT(table_matrix[0])
            for row in range(1, len(table_matrix)):
                table.add_row(table_matrix[row])
            print table
    return table_matrix
