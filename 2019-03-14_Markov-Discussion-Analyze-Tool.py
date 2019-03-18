#!/usr/bin/env python
# coding: utf8

"""

--------------------------------------
| Markov-Discussion-Analysis_Tool.py |
--------------------------------------

A tool, to calculate transition probabilites matrices based on data from a csv.-file.

This tool calculates the following:

- relative transition frequencies,
- absolute transition frequencies,
- expected frequencies,
- z-scores.

The calculated results are saved in a xlsx.-file.

----------------------------------------------------
| v1.06 - date: 2019-03-11 23:25 		Blumenfeld |
----------------------------------------------------

"""


# Import modules from libraries:
import numpy as np
import math
import csv
import xlsxwriter

from sys import stdin, stdout, stderr, argv, exit
from re import *
from os.path import *
from sklearn.preprocessing import normalize
from decimal import Decimal


# Read data from csv.-file into a list, one column with ';' as delimiter:
csv = np.genfromtxt ('2019-03-09_Test.csv', delimiter=";")
a = (csv.tolist())

# Create a list with categories:
categories = [11, 21, 12, 22, 3]


# Calculates the size of the matrix:
# Number of columns (w) and lines (h) in dependence on the number of categories:
w, h = len(categories), len(categories)
 
# Create a new matrix with defined size (w x h) and fill it with zeros:
Matrix = [[0 for x in range(w)] for y in range(h)]

# Define a variable for unvalid values:
ungueltig = 0

i = 0
for x in a:
	
	i = i + 1
	
	j = 0 
	
	if i < len(a):
	
		b = a[i]
		
		# If tuple includes an unvalid value (99), skip this tuple:		
		if x == 99 or b == 99:
			
			# Counts the number of unvalid values:
			if x == 99:
				ungueltig = ungueltig + 1
			pass
		
		else:
		
			# Checks all combinations which begin with k:
			j = 0
			
			for k in categories:
			
				if x == k:
					
					h = 0
					
					# Checks all combinations which end with h:
					for l in categories:
									
						if b == l:
							# Add number to position in matrix:
							Matrix[j][h] = Matrix[j][h] + 1
				
						else:
							h = h + 1
			
				else:
					j = j + 1
					pass
			
			else:
				j = j + 1
				pass
	
	else:
		pass

# Number of scanned values including invalid values (99):
nWerte = len(a)

# Number of valid values excluding invalid values (99)
gueltig = nWerte-ungueltig


# Transform absolute frequencies to relative frequencies in each line in matrix:
Matrix_rel = normalize(Matrix, axis=1, norm='l1')

# Round all elements to 4 digits in the matrix:
Matrix_rel = np.round(Matrix_rel, 4)


# Add a column and a row to get an expanded matrix:
Matrix_erw = [[0 for x in range(w+1)] for y in range(w+1)]


# Create a new column and add to the matrix_erw:
v1 = np.zeros((w, 1))
Matrix_erw = np.c_[Matrix, v1]


# Calculate the sums for each row:
i = 0

for x in range(w):
	
	Matrix_erw [i][w] = np.sum(Matrix[i])
	i = i + 1


# Calculate the sums for each column:
sums_rows = list(np.asarray(Matrix_erw).sum(axis=0))

# Add sums to the expanded matrix:
Matrix_erw = np.vstack((Matrix_erw, sums_rows))

# Convert values from floats to integers:
Matrix_erw = Matrix_erw.astype(int)


# Create an empty matrix for the expected frequencies:
Matrix_exp = [[0 for x in range(w)] for y in range(w)] 


# Fill the matrix with values:
i = 0

for x in range(w):
	
	j = 0
	
	# Berechnung des Wertes je Spalte:
	for x in range(w):
	
		Matrix_exp[i][j] = float(format(sums_rows[j]*np.sum(Matrix[i])/sums_rows[w], '.3f'))
	
		j = j + 1
	
	i = i + 1


# Create an empty matrix for the z-scores:
Matrix_z = [[0 for x in range(w)] for y in range(w)]


# Fill the matrix with values:
i = 0

for x in range(w):
	
	j = 0
	
	for x in range(w):
	
		Matrix_z[i][j] = float(format((Matrix_erw[i][j]-Matrix_exp[i][j])/((Matrix_exp[i][j]*(1-(Matrix_erw[w][j]/Matrix_erw[w][w]))*(1-(Matrix_erw[i][w]/Matrix_erw[w][w])))**0.5), '.4f'))
		
		j = j + 1
	
	i = i + 1
	

#########################################################################################

# Save all results in an xlsx.-file:

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Markov-analyze-results.xlsx')
worksheet1 = workbook.add_worksheet('Observed Frequencies')
worksheet2 = workbook.add_worksheet('Relative Frequencies')
worksheet3 = workbook.add_worksheet('Expected Frequencies')
worksheet4 = workbook.add_worksheet('z-score-Matrix')
worksheet5 = workbook.add_worksheet('Summary')


# Define cell formats:
cell_format_bold = workbook.add_format({'bold': True, 'font_color': 'black'})
cell_format_underline = workbook.add_format({'underline': True, 'font_color': 'black'})

#########################################################################################

# Write absolute frequencies into worksheet1:
row = 0
col = 1
worksheet1.write(0, w+1, 'total', cell_format_bold)
worksheet1.write(w+1, 0, 'total', cell_format_bold)

# Insert labeling:
for item in categories:
	worksheet1.write(row, col, item, cell_format_bold)
	col += 1

col = 0
row = 1

for item in categories:
	worksheet1.write(row, col, item, cell_format_bold)
	row += 1

# Iterate over the data and write it out row by row.
row = 1
col = 1
zeile = 0

for rows in Matrix_erw:
	for item in Matrix_erw[zeile]:
		
		worksheet1.write(row, col, item)
		col += 1
	row += 1
	col = 1
	zeile += 1

#########################################################################################

# Write relative frequencies into worksheet2:

row = 0
col = 1
worksheet2.write(0, w+1, 'total', cell_format_bold)
worksheet2.write(w+1, 0, 'total', cell_format_bold)

# Insert Labeling:
for item in categories:
	worksheet2.write(row, col, item, cell_format_bold)
	col += 1

col = 0
row = 1

for item in categories:
	worksheet2.write(row, col, item, cell_format_bold)
	row += 1

# Iterate over the data and write it out row by row.
row = 1
col = 1
zeile = 0

for rows in Matrix_rel:
	for item in Matrix_rel[zeile]:
		
		worksheet2.write(row, col, item)
		col += 1
	worksheet2.write(row, col, np.round(np.sum(Matrix_rel[zeile]),2))
	row += 1
	col = 1
	zeile += 1

# Insert sums in to the last row and column of the matrix:
col = 1

for item in Matrix_erw[w]:
		
	worksheet2.write(w+1, col, item)
	col += 1

row = 1
for item in Matrix_erw[w]:
		
	worksheet2.write(row, w+1, item)
	row += 1


#########################################################################################

# Write expected frequencies into worksheet3:

row = 0
col = 1
worksheet3.write(0, w+1, 'total', cell_format_bold)
worksheet3.write(w+1, 0, 'total', cell_format_bold)

# Insert Labeling:
for item in categories:
	worksheet3.write(row, col, item, cell_format_bold)
	col += 1

col = 0
row = 1

for item in categories:
	worksheet3.write(row, col, item, cell_format_bold)
	row += 1

# Iterate over the data and write it out row by row.
row = 1
col = 1
zeile = 0

for rows in Matrix_exp:
	for item in Matrix_exp[zeile]:
		
		worksheet3.write(row, col, item)
		col += 1
	row += 1
	col = 1
	zeile += 1

# Insert sums in to the last row and column of the matrix:
col = 1

for item in Matrix_erw[w]:
		
	worksheet3.write(w+1, col, item)
	col += 1

row = 1
for item in Matrix_erw[w]:
		
	worksheet3.write(row, w+1, item)
	row += 1


#########################################################################################

# Write z-scores into worksheet4:

row = 0
col = 1
worksheet4.write(0, w+1, 'total', cell_format_bold)
worksheet4.write(w+1, 0, 'total', cell_format_bold)

# Insert Labeling:
for item in categories:
	worksheet4.write(row, col, item, cell_format_bold)
	col += 1

col = 0
row = 1

for item in categories:
	worksheet4.write(row, col, item, cell_format_bold)
	row += 1

# Iterate over the data and write it out row by row.
row = 1
col = 1
zeile = 0

for rows in Matrix_z:
	for item in Matrix_z[zeile]:
		
		if item > 1.96:
			worksheet4.write(row, col, item, cell_format_bold)
		elif item < -1.96:
			worksheet4.write(row, col, item, cell_format_underline)
		else:
			worksheet4.write(row, col, item)
			
		col += 1
	row += 1
	col = 1
	zeile += 1


# Insert sums in to the last row and column of the matrix:
col = 1

for item in Matrix_erw[w]:
		
	worksheet4.write(w+1, col, item)
	col += 1

row = 1
for item in Matrix_erw[w]:
		
	worksheet4.write(row, w+1, item)
	row += 1

#########################################################################################

# Write summarized results into table:



# Close File:
workbook.close()


# Print the summary:
print("- - - - - - - - - - - - - - - - - - - - -")
print("Number of tested items:" + str(nWerte))
print("Number of unvalid items (99):" + str(ungueltig))
print("- - - - - - - - - - - - - - - - - - - - -")

















	
		
			