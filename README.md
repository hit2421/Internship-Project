# Sales Data Analysis with Python and Excel Integration

## Project Overview

This project performs comprehensive sales data analysis using Python, Excel, and Streamlit. It reads sales data from an Excel file, performs basic data analysis, and displays the results through an interactive dashboard. Users can interact with the data, apply sorting mechanisms, and view various types of charts for an in-depth understanding of sales trends.

## Features

- **Excel Data Integration:** Utilizes the `Openpyxl` and `Pandas` libraries for seamless reading and manipulation of Excel files.
- **Interactive Dashboard:** Built with `Streamlit` and `Plotly`, providing dynamic visualizations and user interactions.
- **Data Visualization:** Supports multiple chart types including bar charts, pie charts, and line graphs.
- **Sorting Mechanisms:** Allows users to sort data based on different criteria for customized analysis.
- **Error Handling:** Includes robust error handling for data loading issues, invalid user inputs, and other potential errors.

## Project Structure


├── main.py                  # The main script to run the application
├── requirements.txt         # List of Python dependencies
├── data                     # Directory to store input Excel files
│   └── sales_data.xlsx      # Example Excel file with sales data
├── README.md                # Project documentation
└── .gitignore               # Git ignore file

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Main.py Overview</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 0;
            background-color: #f4f4f4;
        }
        .container {
            width: 80%;
            margin: auto;
            overflow: hidden;
            padding: 20px;
        }
        h1, h2, h3 {
            color: #333;
        }
        .section {
            background: #fff;
            padding: 15px;
            margin-bottom: 10px;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }
        .section ul {
            list-style: disc;
            margin-left: 20px;
        }
        .section ul li {
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="section">
            <h1>main.py Overview</h1>
            <p>The <code>main.py</code> script is the core of the project, responsible for:</p>
            <ul>
                <li>Initializing the Streamlit dashboard</li>
                <li>Loading and processing the Excel data</li>
                <li>Generating visualizations</li>
                <li>Handling user interactions</li>
            </ul>
        </div>

        <div class="section">
            <h2>Visualization Techniques</h2>
            <ul>
                <li><strong>Bar Chart:</strong> Displays sales performance across different categories.</li>
                <li><strong>Pie Chart:</strong> Shows the distribution of sales across various segments.</li>
                <li><strong>Line Chart:</strong> Tracks sales trends over time.</li>
            </ul>
        </div>

        <div class="section">
            <h2>Sorting Mechanisms</h2>
            <ul>
                <li><strong>Alphabetical Sorting:</strong> Sorts data alphabetically by product name or category.</li>
                <li><strong>Numerical Sorting:</strong> Sorts data based on numerical values such as sales amount or quantity.</li>
                <li><strong>Date Sorting:</strong> Orders data chronologically.</li>
            </ul>
        </div>

        <div class="section">
            <h2>Error Handling</h2>
            <p>The project includes mechanisms to handle:</p>
            <ul>
                <li><strong>File Not Found Error:</strong> Checks if the Excel file is present in the specified directory before loading.</li>
                <li><strong>Data Type Errors:</strong> Validates data types in the Excel file for compatibility with the analysis.</li>
                <li><strong>User Input Validation:</strong> Ensures that user inputs for sorting and chart selection are valid.</li>
            </ul>
        </div>
    </div>
</body>
</html>

