import matplotlib.pyplot as plt; plt.rcdefaults()
import numpy as np
import matplotlib.pyplot as plt
import xlrd

def open_file():
    path = 'C:\Python34\Project\Test1.xlsx'
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
                     color='b',
                     error_kw=error_config,
                     label='Men')
    rects2 = plt.bar(index + bar_width, means_women, bar_width,
                     alpha=opacity,
                     color='r',
                     error_kw=error_config,
                     label='Women')
    plt.ylabel('Read')
    plt.xlabel('Age')
    plt.title('Reading_Thailand')
    plt.xticks(index + bar_width, ('6-14', '15-24', '25-39', '40-59', '60+'))
    plt.legend()
    plt.tight_layout() 
    plt.show()
def main():
    open_file()
main()
    
