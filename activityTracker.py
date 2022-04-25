import glob
import os
import numpy as np
import csv
from pylab import *
import pandas as pd
from scipy import signal
from scipy.fft import rfft, rfftfreq
from scipy.signal import find_peaks
import matplotlib.pyplot as plt
import xlsxwriter

"""Set parent directory"""
#parent_directory = '/Users/nicolelin/Desktop//Fall2021/Duvall/activityTracker-Dong/'

"""designate folder containing data to analyze"""
data_dir = 'RAWOUTPUT/'

"""make a results folder within data_dir"""
path = os.path.join(data_dir, "results")
isExist = os.path.exists(path)
if not isExist:
    os.makedirs(path)

"""load and sort image file names"""
datafiles = glob.glob(os.path.join(data_dir, '*xlsx'))
datafiles.sort(key=lambda x: int(''.join(filter(str.isdigit, x)))) 

"""select wanted sheets from each excel sheet"""
# update the following code with a list of which sheets you need from each excel sheet
    # i.e. excel_1 = [0,1,5] means grab sheets 1, 2, 6 from the first data file
# keep in mind that python indices begin from 0, so sheet 1 will have index 0
# you will have to add/remove to the list when you have more/less data files that you would like to analyze
    # i.e. if you had another data file
        # add excel_5 = [2,3,4,5]
        # add + len(excel_5) to total_length
        # add excel_5 to sheets array
    
total_length = 0
excel_1 = [0,1,2,3,4,5]
excel_2 = [0,1,2,3,4,5]
excel_3 = [0,1,2,3,4,5]
excel_4 = [0,1,2,3,4,5]
excel_5 = [0,1,2,3,4,5]
excel_6 = [0,1,2,3,4,5]
excel_7 = [0,1,2,3,4,5]
excel_8 = [0,1,2,3,4,5]
excel_9 = [0,1,2,3,4,5]
excel_10 = [0,1,2,3,4,5]
excel_11 = [0,1,2,3,4,5]
excel_12 = [0,1,2,3,4,5]
excel_13 = [0,1,2,3,4,5]
excel_14 = [0,1,2,3,4,5]
excel_15 = [0,1,2,3,4,5]
excel_16 = [0,1,2,3,4,5]

total_length = len(excel_1) + len(excel_2) + len(excel_3) + len(excel_4)
total_length = 6*16
sheets = [excel_1, excel_2, excel_3, excel_4, excel_5, excel_6, excel_7, excel_8, excel_9, excel_10, excel_11, excel_12, excel_13, excel_14, excel_15, excel_16]
x = 0
data = []

"""make dataframe for all raw distance_moved for every excel file and nested sheets"""
for file in datafiles:
    print(file)
    #sheet_name = index of sheets starting from 0
    df = pd.read_excel(file, sheet_name = sheets[x], usecols= 'H', header = 36)
    for column in df:
        data.append(df[column])
    x += 1
    
#set normalize to True (activity score) or False (raw distance moved)
normalize = False

#set bin size = 5, 15, or 30 
bin_size = 30

"""This part of the script aims to calculate the periodicity"""
"""The first part is the same procedure, see above for details and comments about calculating act_score"""

x=0
ind_sheet = 0
column = 1

# change name of output file here
excel_name = path + "/raw_bin30.xlsx"
print("writing to " + excel_name + "...")
wb = xlsxwriter.Workbook(excel_name)
worksheet = wb.add_worksheet()
worksheet.write(0, 0, "TS (Time Series)")

#iterating through each of the datafiles in the input folder
for datafile in range(len(datafiles)):
    #iterating through each sheet in each of the datafiles
    for sheet in sheets[x]:
        row = 1
        """This part of the script aims to calculate the periodicity"""
        """The first part is the same procedure, see above for details and comments about calculating act_score"""
        #replace all dashes(-) with NaN
        new_data = data[ind_sheet].replace('-', np.nan).astype('float')
        #reshape using bin_size delineated above
        new_data = new_data.values[: int(np.floor(len(new_data) / bin_size) * bin_size)].reshape((-1, bin_size))
        masked_data = np.ma.masked_array(new_data, np.isnan(new_data))
        #calculate mean of the masked_data
        mean = np.mean(masked_data, axis=1).data
        #normalize to mean
        raw_score = mean / max(mean) * 1000
        act_score = raw_score.astype(int)
        
        if normalize:
            name = "normalized"
            #gives each column name for output excel file
            column_name = datafiles[datafile][19:-5] + "_sheet"+ str(sheet + 1) + "_" + name + "_" + str(bin_size)
            #write column name to excel file -> documentation is (row, column, name)
            worksheet.write(0, column, column_name)
            #finds length of act_score to write in time series (below)
            length = len(act_score)
            #write act_score into excel file
            for score in act_score:
                worksheet.write(row, column, score)
                row +=1
            column += 1
            
        else:
            name = "raw"
            #gives each column name for output excel file
            column_name = datafiles[datafile][13:-5] + "_sheet"+ str(sheet + 1) + "_" + name + "_" + str(bin_size)
            #write column name to excel file -> documentation is (row, column, name)
            worksheet.write(0, column, column_name)
            #finds length of mean data to write in time series (below)
            length = len(mean)
            #write mean into excel file
            for me in mean:
                worksheet.write(row, column, me)
                row +=1
            column += 1
        
        
        #generates time series based on bin_size
        time = 0
        count = 1
        if bin_size == 5:
            for i in range(length):
                worksheet.write(count, 0, time)
                time += round(float(1/12),3)
                count += 1
        #i.e. if bin_size is 15, then iterate "length" (determined above) amount of times and generate the times series by incrementing the time by 0.25 
        elif bin_size == 15:
            for i in range(length):
                worksheet.write(count, 0, time)
                time += 0.25
                count += 1
        else:
            for i in range(length):
                worksheet.write(count, 0, time)
                time += 0.5
                count += 1 
        
        ind_sheet += 1
    x += 1

#close workbook to avoid extra data added
wb.close()
print("finished writing to " + excel_name)

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

#LDD = True if LD_dark, false if only LD
LDD = False

#replace file with correct path
file = '/Users/nicolelin/Desktop/Fall2021/Duvall/activityTracker-Dong/RAWOUTPUT/results/normalized_bin15.xlsx'

#replace pdf name with what you want ur results pdf to be
if LDD:
    pdf = PdfPages('normalized_results_LD_dark.pdf')
else:
    pdf = PdfPages('normalized_results_LD.pdf')

#reading the columns of the excel file and storing them into an array
df = pd.read_excel(file, index_col=0)
cols = df.columns.ravel()

#read in time series from excel file
time_series = pd.read_excel(file, usecols = [0])

#boolean variable for setting ticks on graphs for datasets of different length 
tick = False

#iterates though all columns (total_length = # of columns)
for i in range(total_length):
    
    #find column name of current iteration
    column_name = cols[i]
    
    #consolidated_data = data from current column we are looking at
    consolidated_data = pd.read_excel(file, usecols = [i+1])
    
    #finds length of column, aka number of data values in the column that is not NaN
    length = int(sum(~np.isnan(consolidated_data)))
    
    #if the number of data points in the column is 736, set tick = True (normally data is length 640)
    if length == 736:
        tick = True
        length -= 1
        
    #hr is the exact hour in the Time Series based on the # of data points
    hr = time_series['TS (Time Series)'][length]
    
    #floor finds how many chunks of 24 hours there are in order to make the "looping graph"
    floor = int(hr//24)
    
    #initialize plt.figure()
    figure = plt.figure()
    
    #when floor is 1, that means there are less than 48 hrs of data points, so generate typical matplotlib graph
    if floor == 1:
        plt.title(column_name)
        plt.xlabel('CT/ZT (hr)')
        plt.ylabel('Activity Score')
        plt.plot(time_series, consolidated_data, color="black", linewidth=0.75)
        plt.axvspan(12, 24, facecolor='#9c9c9c', alpha=0.5)
        plt.axvspan(36, 48, facecolor='#9c9c9c', alpha=0.5)
        plt.show()
    else:
        #when missing data points, convert all NaNs to 0
        consolidated_data = consolidated_data.to_numpy()
        consolidated_data[np.isnan(consolidated_data)] = 0
        
        #x,y are the delimeters of the time chunk we want to look at i.e. 0 to 48 hrs, 24 hrs to 72 hrs
        x = 0
        y = 192
        #set subplots, share y-axis, and set figsize = 5x5
        figure, ax = plt.subplots(floor, sharey = True, figsize=(5,5))
        plt.subplots_adjust(hspace=0)
        
        #first plot with title of column_name
        ax[0].plot(time_series[x:y],consolidated_data[x:y], color="black", linewidth=0.75)
        ax[0].set_title(column_name)
        ax[0].axis("off") #remove frame and ticks
        #add gray
        ax[0].axvspan(12, 24, facecolor='#9c9c9c', alpha=0.5)
        ax[0].axvspan(36, 48, facecolor='#9c9c9c', alpha=0.5)
        #increment x, y by 96 to to get to next time chunk
        x += 96
        y += 96
        
        #generates LD for all time chunks in between
        chunk_x1 = 36
        chunk_y1 = 48
        chunk_x2 = 24
        chunk_y2 = 36
        for num in range(1, floor-1): #144 -> 0,48; 24,72; 48,96; 72,120; 96,144
            ax[num].axvspan(chunk_x1, chunk_y1, facecolor='#9c9c9c', alpha=0.5)
            chunk_x1 += 24
            chunk_y1 += 24
            ax[num].axvspan(chunk_x1, chunk_y1, facecolor='#9c9c9c', alpha=0.5)
            if LDD:
                if num == 1:
                    chunk_x2 += 24
                    chunk_y2 += 24
                elif num == 2:
                    chunk_x2 += 24
                    chunk_y2 += 24
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
                else:
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
                    chunk_x2 += 24
                    chunk_y2 += 24
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
            ax[num].plot(time_series[x:y],consolidated_data[x:y], color="black", linewidth=0.75)
            ax[num].axis("off") #remove frame and ticks
            #add gray
            x += 96
            y += 96
        
        #last time chunk 
        ax[num+1].plot(time_series[x:y],consolidated_data[x:y], color="black", linewidth=0.75)
        
        #get rid of frame
        ax[num+1].spines['top'].set_visible(False)
        ax[num+1].spines['right'].set_visible(False)
        ax[num+1].spines['bottom'].set_visible(False)
        ax[num+1].spines['left'].set_visible(False)
        
        #add gray
        if (num+1) == 5:
            ax[num+1].axvspan(132, 144, facecolor='#9c9c9c', alpha=0.5)
            ax[num+1].axvspan(156, 168, facecolor='#9c9c9c', alpha=0.5)

            if LDD: 
                ax[num+1].axvspan(120, 132, facecolor='#d3d3d3', alpha=0.5)
                ax[num+1].axvspan(144, 156, facecolor='#d3d3d3', alpha=0.5)
        else:
            ax[num+1].axvspan(156, 168, facecolor='#9c9c9c', alpha=0.5)
            ax[num+1].axvspan(180, 192, facecolor='#9c9c9c', alpha=0.5)

            if LDD: 
                ax[num+1].axvspan(144, 156, facecolor='#d3d3d3', alpha=0.5)
                ax[num+1].axvspan(168, 180, facecolor='#d3d3d3', alpha=0.5)
        
        #set ticks for 0,12,24,36,48 hrs
        if tick:
            ax[num+1].get_xaxis().set_ticks([144, 156, 168, 180, 192])
        else:
            ax[num+1].get_xaxis().set_ticks([120, 132, 144, 156, 168])
        ax[num+1].set_xticklabels(["0", "12", "24", "36", "48"])
        
        #set labels for graph
        ax[num+1].set_xlabel("CT/ZT (hr)")
        ax[num+1].set_ylabel("Activity Score")
        ax[num+1].get_yaxis().set_ticks([])
        
        # Hide x labels and tick labels for all but bottom plot.
        #for ax in axis.flat:
            #ax.label_outer()
        
        #show graphs
        plt.show()
    #save plot/figure to .pdf    
    pdf.savefig(figure)       
pdf.close()


#bar plot actogram
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

#LDD = True if LD_dark, false if only LD
LDD = True

#replace file with correct path
file = '/Users/nicolelin/Desktop/Fall2021/Duvall/activityTracker-Dong/RAWOUTPUT/results/normalized_bin15.xlsx'

#replace pdf name with what you want ur results pdf to be
if LDD:
    pdf = PdfPages('testbar.pdf')
else:
    pdf = PdfPages('testbar.pdf')

#reading the columns of the excel file and storing them into an array
df = pd.read_excel(file, index_col=0)
cols = df.columns.ravel()

#read in time series from excel file
time_series = pd.read_excel(file, usecols = [0])

#boolean variable for setting ticks on graphs for datasets of different length 
tick = False

#iterates though all columns (total_length = # of columns)
for i in range(total_length):
    
    #find column name of current iteration
    column_name = cols[i]
    
    #consolidated_data = data from current column we are looking at
    consolidated_data = pd.read_excel(file, usecols = [i+1])
    
    #finds length of column, aka number of data values in the column that is not NaN
    length = int(sum(~np.isnan(consolidated_data)))
    
    #if the number of data points in the column is 736, set tick = True (normally data is length 640)
    if length == 736:
        tick = True
        length -= 1
        
    #hr is the exact hour in the Time Series based on the # of data points
    hr = time_series['TS (Time Series)'][length]
    
    ts = []
    time_s = time_series.to_numpy()
    for i in time_s:
        ts.append(i[0])
    
    #floor finds how many chunks of 24 hours there are in order to make the "looping graph"
    floor = int(hr//24)
    
    #initialize plt.figure()
    figure = plt.figure()
    
    #when missing data points, convert all NaNs to 0
    consolidated_data = consolidated_data.to_numpy()
    consolidated_data[np.isnan(consolidated_data)] = 0
    cdata = []
    for i in consolidated_data:
        cdata.append(i[0])
    
    #when floor is 1, that means there are less than 48 hrs of data points, so generate typical matplotlib graph
    if floor == 1:
        plt.title(column_name)
        plt.xlabel('CT/ZT (hr)')
        plt.ylabel('Activity Score')
        plt.bar(ts, cdata, color="black", linewidth=0.75)
        plt.axvspan(12, 24, facecolor='#9c9c9c', alpha=0.5)
        plt.axvspan(36, 48, facecolor='#9c9c9c', alpha=0.5)
        plt.show()
    else:
        #x,y are the delimeters of the time chunk we want to look at i.e. 0 to 48 hrs, 24 hrs to 72 hrs
        x = 0
        y = 192
        #set subplots, share y-axis, and set figsize = 5x5
        figure, ax = plt.subplots(floor, sharey = True, figsize=(5,5))
        plt.subplots_adjust(hspace=0)
        
        #first plot with title of column_name
        ax[0].bar(ts[x:y],cdata[x:y], color="black", linewidth=0.75)
        ax[0].set_title(column_name)
        ax[0].axis("off") #remove frame and ticks
        #add gray
        ax[0].axvspan(12, 24, facecolor='#9c9c9c', alpha=0.5)
        ax[0].axvspan(36, 48, facecolor='#9c9c9c', alpha=0.5)
        #increment x, y by 96 to to get to next time chunk
        x += 96
        y += 96
        
        #generates LD for all time chunks in between
        chunk_x1 = 36
        chunk_y1 = 48
        chunk_x2 = 24
        chunk_y2 = 36
        for num in range(1, floor-1): #144 -> 0,48; 24,72; 48,96; 72,120; 96,144
            ax[num].axvspan(chunk_x1, chunk_y1, facecolor='#9c9c9c', alpha=0.5)
            chunk_x1 += 24
            chunk_y1 += 24
            ax[num].axvspan(chunk_x1, chunk_y1, facecolor='#9c9c9c', alpha=0.5)
            if LDD:
                if num == 1:
                    chunk_x2 += 24
                    chunk_y2 += 24
                elif num == 2:
                    chunk_x2 += 24
                    chunk_y2 += 24
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
                else:
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
                    chunk_x2 += 24
                    chunk_y2 += 24
                    ax[num].axvspan(chunk_x2, chunk_y2, facecolor='#d3d3d3', alpha=0.5)
            ax[num].bar(ts[x:y],cdata[x:y], color="black", linewidth=0.75)
            ax[num].axis("off") #remove frame and ticks
            #add gray
            x += 96
            y += 96
        
        #last time chunk 
        ax[num+1].bar(ts[x:y],cdata[x:y], color="black", linewidth=0.75)
        
        #get rid of frame
        ax[num+1].spines['top'].set_visible(False)
        ax[num+1].spines['right'].set_visible(False)
        ax[num+1].spines['bottom'].set_visible(False)
        ax[num+1].spines['left'].set_visible(False)
        
        #add gray
        if (num+1) == 5:
            ax[num+1].axvspan(132, 144, facecolor='#9c9c9c', alpha=0.5)
            ax[num+1].axvspan(156, 168, facecolor='#9c9c9c', alpha=0.5)

            if LDD: 
                ax[num+1].axvspan(120, 132, facecolor='#d3d3d3', alpha=0.5)
                ax[num+1].axvspan(144, 156, facecolor='#d3d3d3', alpha=0.5)
        else:
            ax[num+1].axvspan(156, 168, facecolor='#9c9c9c', alpha=0.5)
            ax[num+1].axvspan(180, 192, facecolor='#9c9c9c', alpha=0.5)

            if LDD: 
                ax[num+1].axvspan(144, 156, facecolor='#d3d3d3', alpha=0.5)
                ax[num+1].axvspan(168, 180, facecolor='#d3d3d3', alpha=0.5)
        
        #set ticks for 0,12,24,36,48 hrs
        if tick:
            ax[num+1].get_xaxis().set_ticks([144, 156, 168, 180, 192])
        else:
            ax[num+1].get_xaxis().set_ticks([120, 132, 144, 156, 168])
        ax[num+1].set_xticklabels(["0", "12", "24", "36", "48"])
        
        #set labels for graph
        ax[num+1].set_xlabel("CT/ZT (hr)")
        ax[num+1].set_ylabel("Activity Score")
        ax[num+1].get_yaxis().set_ticks([])
        
        # Hide x labels and tick labels for all but bottom plot.
        #for ax in axis.flat:
            #ax.label_outer()
        
        #show graphs
        plt.show()
    #save plot/figure to .pdf    
    pdf.savefig(figure)       
pdf.close()
