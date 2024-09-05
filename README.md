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

 ```

  ├── Adidas Sales Dashboard             
      └── adidas-logo.jpg  # This image file contains the logo for the Adidas Sales Dashboard project. It is typically used for branding and visual representation in the application or documentation.
      └── Adidas.xlsx      # This Excel file contains the sales data used by the application.
      └── app.py           # The main script to run the application
      └── requirement.txt  # List of Python dependencies
  ├── README.md            # Project documentation
   ```

# main.py Overview

The `main.py` script is the core of the project, responsible for:

- Initializing the Streamlit dashboard
- Loading and processing the Excel data
- Generating visualizations
- Handling user interactions

## Visualization Techniques

- **Bar Chart:** Displays sales performance across different categories.
- **Pie Chart:** Shows the distribution of sales across various segments.
- **Line Chart:** Tracks sales trends over time.

## Sorting Mechanisms

- **Alphabetical Sorting:** Sorts data alphabetically by product name or category.
- **Numerical Sorting:** Sorts data based on numerical values such as sales amount or quantity.
- **Date Sorting:** Orders data chronologically.

## Error Handling

The project includes mechanisms to handle:

- **File Not Found Error:** Checks if the Excel file is present in the specified directory before loading.
- **Data Type Errors:** Validates data types in the Excel file for compatibility with the analysis.
- **User Input Validation:** Ensures that user inputs for sorting and chart selection are valid.

## Installation

To set up the project locally, follow these steps:

1. **Clone the repository:**

    ```
    git clone https://github.com/hit2421/Internship-Project.git
    cd Adidas Sales Dashboard
    ```

2. **Install the dependencies:**

    ```
    pip install -r requirement.txt
    ```

3. **Run the application:**

    ```bash
    streamlit run .\app.py
    ```

## Usage

1. Place your Excel data file in the `data` directory.
2. Run the application using Streamlit.
3. Interact with the dashboard to explore various sales metrics, apply sorting, and view different charts.

## Contributing

Contributions are welcome! If you'd like to contribute:
- Fork the repository and use a feature branch.
- Submit a pull request with your proposed changes.

## License

This project is licensed under the [MIT License](LICENSE) - see the LICENSE file for details.

## Acknowledgements

- `Streamlit`
- `Plotly`
- `Pandas`
- `Openpyxl`

