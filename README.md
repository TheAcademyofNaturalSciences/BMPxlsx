# BMPxlsx
[![Build Status](https://travis-ci.org/joemccann/dillinger.svg?branch=master)](https://travis-ci.org/joemccann/dillinger)

The library takes a dictionary of form {Sheet: {Cell: Value}} and updates the specified Excel file accordingly.

This function was created in support of the Model My Watershed Web application WikiWatersheds that models Best Management Practice (BMP) impacts to reducing water quality impacts. (https://modelmywatershed.org/)

https://pypi.org/project/BMPxlsx/

### Example Function Run
```sh
import BMPxlsx
import os

## Function(dataDictionary, FileName)

## Dictionary - {"SHEET": {"CELL": VALUE, "CELL": VALUE}}

datadict = {
'Sheet1': {'D1': 123.4, 'D2': 567.8},
'Sheet2': {'D1': 123.4, 'D2': 567.8},
'Sheet3': {'D1': 123.4, 'D2': 567.8},
}

## Full Path to File

loc = os.getcwd()
fnme = 'test2.xlsx'
file = os.path.join(loc, fnme)

## Run Function

writer = BMPxlsx.Writer(file)
input1 = {'Sheet1': {'D1': 13.4, 'D2': 47.8},
        'Sheet2': {'D1': 23.4, 'D2': 57.8},
        'Sheet3': {'D1': 33.4, 'D2': 67.8},
    }
input2 = {'Sheet1': {'D1': 23.4, 'D2': 57.8},
        'Sheet2': {'D1': 33.4, 'D2': 67.8},
        'Sheet3': {'D1': 43.4, 'D2': 77.8},
    }
writer.write(input1)
writer.close()

```

### Installation

BMPxlsx was written and for Python versions: 3.6, 3.7, 3.8 and 3.9.

```sh
$ pip install BMPxslx
```
