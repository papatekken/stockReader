# Stock data reader

[![Version](https://img.shields.io/badge/version-1.0-orange)]
[![license](https://img.shields.io/github/license/papatekken/stockReader)]
[![Python](https://img.shields.io/badge/Python-3.7.0-blue)]
[![yfinance](https://img.shields.io/badge/yfinance-0.1.54-blue)]
[![openpyxl](https://img.shields.io/badge/openpyxl-4.16.1-blue)]


A Stock data reader which get stock data online and write into the Excel workbook. It was developed in Python, by using the library [openpyxl] and [yfinance]


## About

This system was requested from one of my friends. He provided some criteria and list of stock ticker symbol, so I developed this program, which using yfinance to capture market data, and using openyxl to update the Excel workbook, to fulfill the request.


## Installation

1. Setup [Python](https://www.python.org/) and [GIT](https://git-scm.com/) in runtime environment

2. Install library [yfiance] (https://pypi.org/project/fix-yahoo-finance/) and  [openpyxl] (https://pypi.org/project/openpyxl/)

3. Clone the repository 
    ```
    git clone https://github.com/papatekken/stockReader stockReader
    ```



5. In root directory of 'stockReader', run following command to start the application, when the application finished the run, a new excel workbook is created with data.
	```
	python getStock.py
	```

## License
[MIT](https://github.com/papatekken/StockReader/blob/main/LICENSE)

## Contact
Created by [@papatekken](papatekken@gmail.com) - feel free to contact me!
