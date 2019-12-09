from statistics import *

#calculates, round, and return the mean of the list 
def calcAvg(data):
        
    avg = round(mean(data), 4)
    return avg 

#Find, round, and return the minimum value of the list 
def minData(data):

    minimum = round(min(data), 4)
    return minimum

#Find, round, and return the maximum value of the list 
def maxData(data):

    maximum = round(max(data), 4)
    return maximum

#Calculates and return the range of the list by subtracting both the min and max of the list 
def rangeData(data):

    rangeD = maxData(data) - minData(data)
    return rangeD

#Calculates and return the standard deviation of the list 
def standardDev(data):

    stddev = round(stdev(data), 4)
    return stddev
