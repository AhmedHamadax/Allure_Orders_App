import streamlit as st
import pandas as pd
import numpy as np
import io

# Function to run the script
def run_script(makhzan_file, new_orders_file):
    error_word = []
    error_names = []
    number_of_orders = 0
    error = 0
    # Read the uploaded Excel files
    sys = pd.read_excel(makhzan_file)
    fullfill = pd.read_excel(new_orders_file)
    
    idx = 3
    step = 0
    
    # Loop through each order and process
    for code, name in zip((fullfill['Packages'][3:].astype(str)), (fullfill['Unnamed: 13'][3:].astype(str))):
        if code == 'nan':
            break
        order = fullfill.iloc[idx]['Unnamed: 31']
        order_splitted = (order.split(' '))
        count1 = order_splitted[0]
        product1 = order_splitted[2]
        try:
            count2 = order_splitted[3]
            product2 = order_splitted[5]
        except:
            count2 = 0
            product2 = ''
        try:
            count3 = order_splitted[6]
            product3 = order_splitted[8]
        except:
            product3 = ''
            count3 = 0
        all_products = [prod for prod in [product1, product2, product3] if (prod != '')]
        all_counts = [c for c in [count1, count2, count3] if (c != 0)]
        sys.loc[idx - 3 + step, 'TN'] = code
        number_of_orders += 1
        sys.loc[idx - 3 + step, 'CST Name'] = name

        # Update SKU and Quantity based on product type
        for c, prod in zip(all_counts, all_products):
            if prod == 'بشرة':
                sys.loc[idx - 3 + step, 'SKU'] = 'Allure15948'
                sys.loc[idx - 3 + step, 'Quantity'] = c
            elif prod == 'عين':
                sys.loc[idx - 3 + step, 'SKU'] = 'Allure12345'
                sys.loc[idx - 3 + step, 'Quantity'] = c
            elif prod == 'شعر':
                sys.loc[idx - 3 + step, 'SKU'] = 'Allure67890'
                sys.loc[idx - 3 + step, 'Quantity'] = c
            elif prod == 'اسكرين':
                sys.loc[idx - 3 + step, 'SKU'] = 'Allure35724'
                sys.loc[idx - 3 + step, 'Quantity'] = c
            else:
                error_word.append(prod)
                error_names.append(name)
                error = 1
            step += 1

        idx += 1

    # Streamlit UI to show success or error message
    if error == 1:
        st.error(f"undefined word found: ({error_word} for {error_names} orders)")
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
    
    # Convert the final dataframe to a BytesIO object to enable download
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
