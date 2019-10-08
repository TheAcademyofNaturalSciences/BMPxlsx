import BMPxlsx
import os


# Example Function Run
# Function(dataDictionary, FileName)

# DICTIONARY {"SHEET": {"CELL": VALUE, "CELL": VALUE}}
datadict = {'Sheet1': {'D1': 123.4, 'D2': 567.8},
            'Sheet2': {'D1': 123.4, 'D2': 567.8},
            'Sheet3': {'D1': 123.4, 'D2': 567.8},
            }

# FULL PATH TO FILE
loc = os.getcwd()
fnme = 'test2.xlsx'
file = os.path.join(loc, fnme)

# Function Methods
writer = BMPxlsx.Writer(file)
input1 = {'Sheet1': {'D1': 13.4, 'D2': 47.8},
        'Sheet2': {'D1': 23.4, 'D2': 57.8},
        'Sheet3': {'D1': 33.4, 'D2': 67.8},
    }
input2 = {'Sheet1': {'D1': 23.4, 'D2': 57.8},
        'Sheet2': {'D1': 33.4, 'D2': 67.8},
        'Sheet3': {'D1': 43.4, 'D2': 77.8},
        'Sheet4': {'D2': 45.0}
    }
writer.write(input1)
writer.close()

