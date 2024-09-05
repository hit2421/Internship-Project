import streamlit as st
import pandas as pd
import datetime
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
from wordcloud import WordCloud, STOPWORDS
import networkx as nx
import matplotlib.pyplot as plt

# Load the Excel file
try:
    df = pd.read_excel("Adidas.xlsx")
except FileNotFoundError:
    st.error("Error: The file 'Adidas.xlsx' was not found.")
    st.stop()
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()

# Set page layout
st.set_page_config(layout="wide")

# Custom CSS for reducing top padding
st.markdown('<style>div.block-container{padding-top:2.1rem;}</style>', unsafe_allow_html=True)

# Load and display the logo
try:
    image = Image.open('adidas-logo.jpg')
except FileNotFoundError:
    st.error("Error: The logo file 'adidas-logo.jpg' was not found.")
    st.stop()
except Exception as e:
    st.error(f"Error loading image: {e}")
    st.stop()

col1, col2 = st.columns([0.1, 0.9])

with col1:
    st.image(image, width=100)

html_title = """
<style>
    .title-test {
        font-size: 26px;
        font-weight: bold;
        padding: 5px;
        border-radius: 6px;
        text-align: center;
    }
</style>
<h1 class="title-test">Adidas Interactive Sales Dashboard</h1>"""

with col2:
    st.markdown(html_title, unsafe_allow_html=True)

# Sorting options
sort_order = st.radio("Sort Data", options=["Ascending", "Descending"])

# Helper function for sorting data
def sort_data(df, column, order):
    return df.sort_values(by=column, ascending=(order == "Ascending"))

# Display the last updated date
# sourcery skip: identity-comprehension, remove-unnecessary-cast
col3, col4, col5 = st.columns([0.1, 0.45, 0.45])

with col3:
    box_date = str(datetime.datetime.now().strftime("%d %B %Y"))
    st.write(f"Last updated by: \n {box_date}")

try:
    # Total Sales by Retailer
    sorted_df = sort_data(df, "TotalSales", sort_order)
    with col4:
        fig = px.bar(sorted_df, x="Retailer", y="TotalSales", labels={"TotalSales": "Total Sales {$}"},
                     title="Total Sales by Retailer", hover_data=["TotalSales"],
                     template="gridon", height=500, color="Retailer", 
                     color_discrete_map={product: color for product, color in zip(sorted_df["Retailer"].unique(), px.colors.qualitative.Antique)})
        st.plotly_chart(fig, use_container_width=True)

    _, view1, dwn1, view2, dwn2 = st.columns([0.15, 0.20, 0.20, 0.20, 0.20])

    with view1:
        expander = st.expander("Retailer wise Sales")
        data = sorted_df[["Retailer", "TotalSales"]].groupby(by="Retailer")["TotalSales"].sum()
        expander.write(data)

    with dwn1:
        st.download_button("Get Data", data=data.to_csv().encode("utf-8"),
                           file_name="RetailerSales.csv", mime="text/csv")

    # Total Sales Over Time
    df["Month_Year"] = df["InvoiceDate"].dt.strftime("%b'%y")
    result = df.groupby(by="Month_Year")["TotalSales"].sum().reset_index()
    sorted_result = sort_data(result, "TotalSales", sort_order)

    with col5:
        fig1 = px.line(sorted_result, x="Month_Year", y="TotalSales", title="Total Sales Over Time",
                       template="gridon")
        st.plotly_chart(fig1, use_container_width=True)

    with view2:
        expander = st.expander("Monthly Sales")
        expander.write(sorted_result)

    with dwn2:
        st.download_button("Get Data", data=sorted_result.to_csv().encode("utf-8"),
                           file_name="MonthlySales.csv", mime="text/csv")
except Exception as e:
    st.error(f"An error occurred: {e}")

st.divider()

try:
    # Total Sales and Units Sold by State
    result1 = df.groupby(by="State")[["TotalSales", "UnitsSold"]].sum().reset_index()
    sorted_result1 = sort_data(result1, "TotalSales", sort_order)

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(x=sorted_result1["State"], y=sorted_result1["TotalSales"], name="Total Sales",
                          marker_color=[px.colors.qualitative.Antique[i % len(px.colors.qualitative.Antique)] 
                                        for i in range(len(sorted_result1["State"]))]))
    fig3.add_trace(go.Scatter(x=sorted_result1["State"], y=sorted_result1["UnitsSold"], mode="lines",
                              name="Units Sold", yaxis="y2"))
    fig3.update_layout(
        title="Total Sales and Units Sold by State",
        xaxis=dict(title="State"),
        yaxis=dict(title="Total Sales", showgrid=False),
        yaxis2=dict(title="Units Sold", overlaying="y", side="right"),
        template="gridon",
        legend=dict(x=1, y=1)
    )

    _, col6 = st.columns([0.1, 1])
    with col6:
        st.plotly_chart(fig3, use_container_width=True)

    _, view3, dwn3 = st.columns([0.5, 0.45, 0.45])
    with view3:
        expander = st.expander("View Data for Sales by Units Sold")
        expander.write(sorted_result1)

    with dwn3:
        st.download_button("Get Data", data=sorted_result1.to_csv().encode("utf-8"),
                           file_name="Sales_by_UnitsSold.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check the column names in your dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. Please check the data types in your dataset.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Total Sales by Region and City in Treemap
    treemap = df[["Region", "City", "TotalSales"]].groupby(by=["Region", "City"])["TotalSales"].sum().reset_index()
    sorted_treemap = sort_data(treemap, "TotalSales", sort_order)

    def format_sales(value):
        if value >= 0:
            return '{:.2f} Lakh'.format(value / 1_00_000)

    sorted_treemap["TotalSales (Formatted)"] = sorted_treemap["TotalSales"].apply(format_sales)
    fig4 = px.treemap(sorted_treemap, path=["Region", "City"], values="TotalSales",
                      hover_name="TotalSales (Formatted)",
                      hover_data=["TotalSales (Formatted)"],
                      color="City", height=700, width=600)

    fig4.update_traces(textinfo="label+value")

    _, col7 = st.columns([0.1, 1])
    with col7:
        st.subheader("👉Total Sales by Region and City in Treemap")
        st.plotly_chart(fig4, use_container_width=True)

    _, view4, dwn4 = st.columns([0.5, 0.45, 0.45])
    with view4:
        result2 = sorted_treemap
        expander = st.expander("View data for Total Sales by Region and City")
        expander.write(result2)

    with dwn4:
        st.download_button("Get Data", data=result2.to_csv().encode("utf-8"),
                           file_name="Sales_by_Region.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'Region', 'City', or 'TotalSales' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or the format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Total Sales by City in Piechart
    piechart = df.groupby(by=['City'])['TotalSales'].sum().reset_index()
    sorted_piechart = sort_data(piechart, "TotalSales", sort_order)

    fig5 = px.pie(sorted_piechart, values='TotalSales', names='City')
    fig5.update_layout(margin=dict(l=0, r=0, t=50, b=0))

    _, col8 = st.columns([0.1, 1])
    with col8:
        st.subheader("👉Total Sales by City in Piechart")
        st.plotly_chart(fig5, use_container_width=True)

    _, view6, dwn6 = st.columns([0.5, 0.45, 0.45])
    with view6:
        result3 = sorted_piechart
        expander = st.expander("View data for Total Sales by City")
        expander.write(result3)

    with dwn6:
        st.download_button("Get Data", data=sorted_piechart.to_csv().encode("utf-8"),
                           file_name="Sales_by_City.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'City' or 'TotalSales' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or the format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Sales Share by Product in Donut Chart
    product_sales = df[["Product", "TotalSales"]].groupby(by="Product").sum().reset_index()
    sorted_product_sales = sort_data(product_sales, "TotalSales", sort_order)

    fig9 = px.pie(sorted_product_sales, values='TotalSales', names='Product', hole=.3,
                  hover_data=['TotalSales'], labels={'TotalSales':'Sales'})

    fig9.update_traces(textinfo='percent+label')

    _, col9 = st.columns([0.1, 1])
    with col9:
        st.subheader(" Sales Share by Product")
        st.plotly_chart(fig9, use_container_width=True)

    _, view7, dwn7 = st.columns([0.5, 0.45, 0.45])
    with view7:
        result4 = sorted_product_sales
        expander = st.expander("View data for Sales Share by Product")
        expander.write(result4)

    with dwn7:
        st.download_button("Get Data", data=sorted_product_sales.to_csv().encode("utf-8"),
                           file_name="Sales_Share_by_Product.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'Product' or 'TotalSales' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or the format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Correlation Heatmap
    corr_cols = ["PriceperUnit", "UnitsSold", "TotalSales", "OperatingProfit", "OperatingMargin"]
    corr = df[corr_cols].corr()

    fig10 = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale='RdBu_r')

    _, col10 = st.columns([0.1, 1])
    with col10:
        st.subheader(" Correlation Heatmap of Sales Data")
        st.plotly_chart(fig10, use_container_width=True)

    _, view8, dwn8 = st.columns([0.5, 0.45, 0.45])
    with view8:
        result5 = df[corr_cols].corr()
        expander = st.expander("View data for Correlation Heatmap")
        expander.write(result5)

    with dwn8:
        st.download_button("Get Data", data=corr.to_csv().encode("utf-8"),
                           file_name="Correlation_Heatmap_of_SalesData.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the required columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")


st.divider()

try:
    # Scatterplot: Units Sold vs Price per Unit
    sorted_scatter = sort_data(df, "UnitsSold", sort_order)
    fig11 = px.scatter(sorted_scatter, x="PriceperUnit", y="UnitsSold")

    _, col11 = st.columns([0.1, 1])
    with col11:
        st.subheader(" Units Sold vs Price per Unit")
        st.plotly_chart(fig11, use_container_width=True)

    _, view9, dwn9 = st.columns([0.5, 0.45, 0.45])
    with view9:
        result6 = sorted_scatter[["PriceperUnit", "UnitsSold"]]
        expander = st.expander("View data for Units Sold vs Price per Unit")
        expander.write(result6)

    with dwn9:
        st.download_button("Get Data", data=result6.to_csv().encode("utf-8"),
                           file_name="Units_Sold_vs_Price_per_Unit.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'PriceperUnit' or 'UnitsSold' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Histogram: Distribution of Operating Profit by Product
    sorted_hist = sort_data(df, "OperatingProfit", sort_order)
    fig12 = px.histogram(sorted_hist, x="Product", y="OperatingProfit", nbins=50, 
                         color="Product", color_discrete_map={product: color for product, color in zip(sorted_hist["Product"].unique(), px.colors.qualitative.Antique)})

    _, col12 = st.columns([0.1, 1])
    with col12:
        st.subheader("Distribution of Operating Profit by Product")
        st.plotly_chart(fig12, use_container_width=True)

    _, view10, dwn10 = st.columns([0.5, 0.45, 0.45])
    with view10:
        result7 = sorted_hist[["Product", "OperatingProfit"]]
        expander = st.expander("View data for Distribution of Operating Profit by Product")
        expander.write(result7)

    with dwn10:
        st.download_button("Get Data", data=result7.to_csv().encode("utf-8"),
                           file_name="Distribution_of_Operating_Profit_by_Product.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'Product' or 'OperatingProfit' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")


st.divider()

try:
    # Box plot
    sorted_box_plot = sort_data(df, "TotalSales", sort_order)
    fig13 = px.box(sorted_box_plot, x="Retailer", y="TotalSales", title="Boxplot of Total Sales by Retailer",
                   color="Retailer", color_discrete_map={product: color for product, color in zip(sorted_box_plot["Retailer"].unique(), px.colors.qualitative.Antique)})
    st.plotly_chart(fig13, use_container_width=True)

    st.divider()

    # Bubble chart
    sorted_bubble_chart = sort_data(df, "Region", sort_order)
    fig14 = px.scatter(sorted_bubble_chart, x="Product", y="Region", size="UnitsSold", color="Product",
                       title="Bubble Chart of Units Sold by Product and Region")
    st.plotly_chart(fig14, use_container_width=True)

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the required columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")


st.divider()


try:
    # Area Chart: Total Sales by Month
    df["Month_Year"] = df["InvoiceDate"].dt.strftime("%b'%y")
    result = df.groupby(by="Month_Year")["TotalSales"].sum().reset_index()
    sorted_result = sort_data(result, "TotalSales", sort_order)

    fig15 = px.area(sorted_result, x="Month_Year", y="TotalSales")

    _, col13 = st.columns([0.1, 1])
    with col13:
        st.subheader("Area Chart of Total Sales by Month")
        st.plotly_chart(fig15, use_container_width=True)

    _, view12, dwn12 = st.columns([0.5, 0.45, 0.45])
    with view12:
        result8 = sorted_result
        expander = st.expander("View data for Area Chart of Total Sales by Month")
        expander.write(result8)

    with dwn12:
        st.download_button("Get Data", data=result8.to_csv().encode("utf-8"),
                           file_name="Area_Chart_of_Total_Sales_by_Month.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'InvoiceDate' or 'TotalSales' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")


st.divider()

try:
    # Create a highlight table
    highlight_table = sorted_result.style.format({"TotalSales": "${:,.2f}M"}).highlight_max(axis=0, props="background-color:cyan;color:black;")
    fig16 = go.Figure(data=[go.Table(
        header=dict(values=["Month_Year", "TotalSales"],
                    line_color='darkslategray',
                    fill_color='lightskyblue',
                    align='left'),
        cells=dict(values=[sorted_result["Month_Year"], sorted_result["TotalSales"]],
                   line_color='darkslategray',
                   fill_color='lightcyan',
                   align='left'))
    ])
    
    _, col14 = st.columns([0.1, 1])
    with col14:
        st.subheader("Highlight Table of Total Sales by Month")
        st.table(highlight_table)

    _, view13, dwn13 = st.columns([0.5, 0.45, 0.45])
    with view13:
        result9 = sorted_result
        expander = st.expander("View data for Highlight Table of Total Sales by Month") 
        expander.write(result9)

    with dwn13:
        st.download_button("Get Data", data=result9.to_csv().encode("utf-8"),
                           file_name="Highlight_Table_of_Total_Sales_by_Month.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'Month_Year' or 'TotalSales' columns exist in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")


st.divider()

try:
    # Create a word cloud 
    wordcloud_df = df["Product"].value_counts().reset_index()
    wordcloud_df.columns = ["Product", "Frequency"]

    # Apply sorting based on the selected order
    wordcloud_df = sort_data(wordcloud_df, "Frequency", sort_order)

    # Generate the word cloud
    wordcloud = WordCloud(width=800, height=400, random_state=21, max_font_size=110).generate_from_frequencies(wordcloud_df.set_index("Product")["Frequency"])

    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation="bilinear")
    plt.axis("off")
    plt.title(f"Word Cloud of Products - {sort_order} Order")

    _, col15 = st.columns([0.1, 1])
    with col15:
        st.subheader("Word Cloud of Products")
        st.pyplot(plt)

    _, view14, dwn14 = st.columns([0.5, 0.45, 0.45])
    with view14:
        expander = st.expander(f"View data for Word Cloud of Products")
        expander.write(wordcloud_df)

    with dwn14:
        st.download_button("Get Data", data=wordcloud_df.to_csv().encode("utf-8"),
                           file_name=f"WordCloud_of_Products_{sort_order}.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if the 'Product' column exists in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data types or format of the values.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

try:
    # Define the nodes and links for the Sankey diagram
    nodes = dict(
        label = ["Retailer", "Region", "City", "Product", "Sales"],
        color = "blue",
        pad = 15,
        thickness = 20,
        line = dict(color = "black", width = 0.5)
    )

    links = dict(
        source = [0, 1, 2, 3], # indices correspond to labels, eg A1, A2, etc.
        target = [4, 4, 4, 4],
        value = [df["TotalSales"].sum(), df["TotalSales"].sum(), df["TotalSales"].sum(), df["TotalSales"].sum()]
    )

    # Create the Sankey diagram
    fig = go.Figure(data=[go.Sankey(
        node = nodes,
        link = links
    )])

    # Update the layout
    fig.update_layout(title_text="Sankey Diagram of Adidas Sales", font_size=16)
    fig.show()

    # Display the Sankey diagram in the Streamlit app
    _, col16 = st.columns([0.1, 1])
    with col16:
        st.plotly_chart(fig, use_container_width=True)

    # Add a download button for the Sankey diagram data
    _, view15, dwn15 = st.columns([0.5, 0.45, 0.45])
    with view15:
        expander = st.expander("View data for Sankey Diagram of Adidas Sales")
        expander.write(df[["Retailer", "Region", "City", "Product", "TotalSales"]])

    with dwn15:
        st.download_button("Get Data", data=df[["Retailer", "Region", "City", "Product", "TotalSales"]].to_csv().encode("utf-8"),
                           file_name="Sankey_Diagram_of_Adidas_Sales.csv", mime="text/csv")

    st.divider()

    # Add a section to view and download raw sales data
    _, view15, dwn15 = st.columns([0.5, 0.45, 0.45])
    with view15:
        expander = st.expander("View Sales Raw Data")
        expander.write(df)

    with dwn15:
        st.download_button("Get Raw Data", data=df.to_csv().encode("utf-8"),
                           file_name="SalesRawData.csv", mime="text/csv")

except KeyError as e:
    st.error(f"Key error: {e}. Please check if all necessary columns are present in the dataset.")
except ValueError as e:
    st.error(f"Value error: {e}. There might be an issue with the data values or their format.")
except Exception as e:
    st.error(f"An unexpected error occurred: {e}")

st.divider()

# Display basic statistics
st.header("Basic Statistics")
st.write(df.describe())


