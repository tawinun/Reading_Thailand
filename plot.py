
import matplotlib.pyplot as plt; plt.rcdefaults()
import numpy as np
import matplotlib.pyplot as plt
import xlrd

def central():
    bar_chart_central()
    bar_chart_c7()
    bar_chart_c9()
def north():
    bar_chart_north()
    bar_chart_n7()
    bar_chart_n9()
def northeast():
    bar_chart_northeast()
    bar_chart_ne7()
    bar_chart_ne9()
def south():
    bar_chart_south()
    bar_chart_s7()
    bar_chart_s9()
def bar_chart_central():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 5
    means_men = (int(data_set[6][3]),int(data_set[6][4]),int(data_set[6][5]),\
                 int(data_set[6][6]),int(data_set[6][7]))
    means_women = (int(data_set[7][3]),int(data_set[7][4]),int(data_set[7][5]),\
                   int(data_set[7][6]),int(data_set[7][7]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='#3300FF',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='#FF0066',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Age')
    plt.title('Reading_Thailand Central')
    plt.xticks(index + bar_width, ('6-14', '15-24', '25-39', '40-59', '60+'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_c7():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 8
    means_men = (int(data_set[49][3]),int(data_set[49][4]),int(data_set[49][5]),\
                 int(data_set[49][6]),int(data_set[49][7]),int(data_set[49][8]),\
                 int(data_set[49][9]),int(data_set[49][10]))
    means_women = (int(data_set[50][3]),int(data_set[50][4]),int(data_set[50][5]),\
                 int(data_set[50][6]),int(data_set[50][7]),int(data_set[50][8]),\
                 int(data_set[50][9]),int(data_set[50][10]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Reason')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e','f','g','h'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_c9():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 4
    means_men = (int(data_set[97][3]),int(data_set[97][4]),int(data_set[97][5]),\
                 int(data_set[97][6]))
    means_women = (int(data_set[98][3]),int(data_set[98][4]),int(data_set[98][5]),\
                 int(data_set[98][6]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Time')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def bar_chart_north():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 5
    means_men = (int(data_set[16][3]),int(data_set[16][4]),int(data_set[16][5]),\
                 int(data_set[16][6]),int(data_set[16][7]))
    means_women = (int(data_set[17][3]),int(data_set[17][4]),int(data_set[17][5]),\
                   int(data_set[17][6]),int(data_set[17][7]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='#CCFF66',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='#66CC66',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Age')
    plt.title('Reading_Thailand North')
    plt.xticks(index + bar_width, ('6-14', '15-24', '25-39', '40-59', '60+'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_n7():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 8
    means_men = (int(data_set[61][3]),int(data_set[61][4]),int(data_set[61][5]),\
                 int(data_set[61][6]),int(data_set[61][7]),int(data_set[61][8]),\
                 int(data_set[61][9]),int(data_set[61][10]))
    means_women = (int(data_set[62][3]),int(data_set[62][4]),int(data_set[62][5]),\
                 int(data_set[62][6]),int(data_set[62][7]),int(data_set[62][8]),\
                 int(data_set[62][9]),int(data_set[62][10]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Reason')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e','f','g','h'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_n9():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 4
    means_men = (int(data_set[108][3]),int(data_set[108][4]),int(data_set[108][5]),\
                 int(data_set[108][6]))
    means_women = (int(data_set[109][3]),int(data_set[109][4]),int(data_set[109][5]),\
                 int(data_set[109][6]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Time')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e'))
    plt.legend()
    plt.tight_layout() 
    plt.show()    
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def bar_chart_northeast():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 5
    means_men = (int(data_set[26][3]),int(data_set[26][4]),int(data_set[26][5]),\
                 int(data_set[26][6]),int(data_set[26][7]))
    means_women = (int(data_set[27][3]),int(data_set[27][4]),int(data_set[27][5]),\
                   int(data_set[27][6]),int(data_set[27][7]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='#FF9933',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='#FF3333',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Age')
    plt.title('Reading_Thailand NorthEast')
    plt.xticks(index + bar_width, ('6-14', '15-24', '25-39', '40-59', '60+'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_ne7():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 8
    means_men = (int(data_set[74][3]),int(data_set[74][4]),int(data_set[74][5]),\
                 int(data_set[74][6]),int(data_set[74][7]),int(data_set[74][8]),\
                 int(data_set[74][9]),int(data_set[74][10]))
    means_women = (int(data_set[75][3]),int(data_set[75][4]),int(data_set[75][5]),\
                 int(data_set[75][6]),int(data_set[75][7]),int(data_set[75][8]),\
                 int(data_set[75][9]),int(data_set[75][10]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Reason')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e','f','g','h'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_ne9():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 4
    means_men = (int(data_set[119][3]),int(data_set[119][4]),int(data_set[119][5]),\
                 int(data_set[119][6]))
    means_women = (int(data_set[120][3]),int(data_set[120][4]),int(data_set[120][5]),\
                 int(data_set[120][6]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Time')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def bar_chart_south():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 4
    means_men = (int(data_set[36][3]),int(data_set[36][4]),int(data_set[36][5]),\
                 int(data_set[36][6]))
    means_women = (int(data_set[37][3]),int(data_set[37][4]),int(data_set[37][5]),\
                   int(data_set[37][6]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='#CD853F',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='#EECFA1',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Age')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('6-14', '15-24', '25-39', '40-59', '60+'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_s7():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 8
    means_men = (int(data_set[86][3]),int(data_set[86][4]),int(data_set[86][5]),\
                 int(data_set[86][6]),int(data_set[86][7]),int(data_set[86][8]),\
                 int(data_set[86][9]),int(data_set[86][10]))
    means_women = (int(data_set[87][3]),int(data_set[87][4]),int(data_set[87][5]),\
                 int(data_set[87][6]),int(data_set[87][7]),int(data_set[87][8]),\
                 int(data_set[87][9]),int(data_set[87][10]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Reason')
    plt.title('Reading_Thailand South')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e','f','g','h'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def bar_chart_s9():
    path = 'C:\Python34\Project\data.xlsx'
    workbook = xlrd.open_workbook(path) 
    worksheet = workbook.sheet_by_index(0) 
    data_set = [[worksheet.cell_value(row,col) for col in range(worksheet.ncols)] for
                row in range(worksheet.nrows)]
    n_groups = 4
    means_men = (int(data_set[130][3]),int(data_set[130][4]),int(data_set[130][5]),\
                 int(data_set[130][6]))
    means_women = (int(data_set[131][3]),int(data_set[131][4]),int(data_set[131][5]),\
                 int(data_set[131][6]))
    fig, ax = plt.subplots()
    index = np.arange(n_groups)
    bar_width = 0.35
    opacity = 0.4
    error_config = {'ecolor': '0.3'}
    rects1 = plt.bar(index, means_men, bar_width,
                     alpha=opacity,
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Time')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('a', 'b', 'c', 'd', 'e'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#    
def pie_chart():
    # Data to plot
    labels = 'Central', 'North', 'NortEast', 'South'
    sizes = [17645244.9931999, 10818759.0109, 17404374.0166001, 7862282.60390006]
    colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
    explode = (0.1, 0, 0, 0)  # explode 1st slice 
    # Plot
    plt.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
 
    plt.axis('equal')
    plt.title('Total of Reading_Thailand')
    plt.legend()
    plt.show()
def main():
    central()
    north()
    northeast()
    south()
main()
