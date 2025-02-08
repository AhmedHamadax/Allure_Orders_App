import streamlit as st
import pandas as pd
import numpy as np
import io

# Function to run the script
def run_script(makhzan_file, new_orders_file):
    # Read fresh copies of the uploaded Excel files
    sys = pd.read_excel(makhzan_file).copy()
    fullfill = pd.read_excel(new_orders_file).copy()
    
    # Reset variables for a clean state
    error_word = []
    error_names = []
    number_of_orders = 0
    error = 0
    idx = 3
    step = 0

    # Loop through each order and process
    for code, name in zip((fullfill['Packages'][3:].astype(str)), (fullfill['Unnamed: 13'][3:].astype(str))):
        if code == 'nan':
            break
        order = fullfill.iloc[idx]['Unnamed: 31']
        order_splitted = order.split(' ')
        
        count1, product1 = order_splitted[0], order_splitted[2]
        count2, product2, count3, product3 = 0, '', 0, ''
        
        try:
            count2, product2 = order_splitted[3], order_splitted[5]
        except IndexError:
            pass
        try:
            count3, product3 = order_splitted[6], order_splitted[8]
        except IndexError:
            pass

        all_products = [prod for prod in [product1, product2, product3] if prod]
        all_counts = [c for c in [count1, count2, count3] if c]

        sys.loc[idx - 3 + step, 'TN'] = code
        sys.loc[idx - 3 + step, 'CST Name'] = name
        number_of_orders += 1

        # Update SKU and Quantity based on product type
        for c, prod in zip(all_counts, all_products):
            sku_map = {
                'بشرة': 'Allure15948',
                'عين': 'Allure12345',
                'شعر': 'Allure67890',
                'اسكرين': 'Allure35724'
            }
            if prod in sku_map:
                sys.loc[idx - 3 + step, 'SKU'] = sku_map[prod]
                sys.loc[idx - 3 + step, 'Quantity'] = c
            else:
                error_word.append(prod)
                error_names.append(name)
                error = 1
            step += 1

        idx += 1

    # Streamlit UI to show success or error message
    if error == 1:
        st.error(f"Undefined words found: {error_word} for {error_names} orders")
    else:
        st.success(f"{number_of_orders} orders successfully assigned")

    return sys

# Streamlit UI setup
st.title("Allure Orders Generator")
st.write("Upload the Excel files below to run the script:")

# File uploaders for the two Excel files
makhzan_file = st.file_uploader("Upload Makhzan Excel File", type=["xlsx"])
new_orders_file = st.file_uploader("Upload New Orders Excel File", type=["xlsx"])

# Button to trigger the script if files are uploaded
if st.button("Run Script") and makhzan_file is not None and new_orders_file is not None:
    final_df = run_script(makhzan_file, new_orders_file)

    if final_df is not None:
        # Convert the DataFrame to Excel and then to BytesIO
        output = io.BytesIO()
        final_df.to_excel(output, index=False)
        output.seek(0)

        # Provide a download button for the user
        st.download_button(
            label="Download Updated Excel File",
            data=output,
            file_name="Allure_Updated_values.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
