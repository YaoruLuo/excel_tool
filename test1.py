import streamlit as st
import pandas as pd
import io

# Custom CSS for improved styling
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


# Function to perform the operation on data
def perform_operation(data, operation, value):
    if operation == "Multiply":
        return data * value
    elif operation == "Add":
        return data + value
    elif operation == "Subtract":
        return data - value
    elif operation == "Divide":
        return data / value
    elif operation == "sum":
        return data.sum()
    elif operation == "Average":
        return data.mean()
    else:
        return data

local_css("style.css")  # Assuming a style.css file is in the same directory

# Page Title and Layout
st.set_page_config(page_title="Excel Processor", layout="wide")

st.markdown("""
<div class="header">
    <h1>Excel Processor</h1>
    <p>Upload and process Excel files with ease</p>
</div>
""", unsafe_allow_html=True)

# Upload Section
st.markdown("""
<div class="upload-section">
    <h2>Step 1: Upload Your Excel File</h2>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Sheet Selection Section
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        sheet = st.sidebar.selectbox("Select a sheet to process", sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=sheet)

        # Data Display Section
        st.markdown("""
        <div class="preview-section">
            <h2>Step 2: Preview Your Data</h2>
        </div>
        """, unsafe_allow_html=True)
        st.dataframe(df)

        # Operations Section
        st.markdown("""
        <div class="operation-section">
            <h2>Step 3: Define and Apply Operations</h2>
        </div>
        """, unsafe_allow_html=True)

        operations = []
        results = []  # Store results for intermediate calculations

        with st.expander("Add Operations", expanded=True):
            operation_type = st.selectbox("Select an operation:", ["Multiply", "Add", "Subtract", "Divide", "Sum", "Average"])
            axis = st.radio("Choose the axis:", ["Column", "Row", "Multiple Columns", "Multiple Rows"], horizontal=True)

            if operation_type in ["Multiply", "Add", "Subtract", "Divide"]:
                if axis == "Column":
                    column = st.selectbox("Select a column to process", df.columns, key="col_select")
                    value = st.number_input("Enter the value:", value=1.0, key="value_input")
                    if st.button("Apply Operation to Column"):
                        df[column] = perform_operation(df[column], operation_type, value)
                        st.write(f"Result for column {column}:")
                        st.dataframe(df[column])
                        results.append((f"Column {column}", df[column].copy()))

                elif axis == "Row":
                    row = st.number_input("Enter the row index:", min_value=0, max_value=len(df)-1, step=1, key="row_select")
                    value = st.number_input("Enter the value:", value=1.0, key="value_input_row")
                    if st.button("Apply Operation to Row"):
                        df.iloc[row] = perform_operation(df.iloc[row], operation_type, value)
                        st.write(f"Result for row {row}:")
                        st.dataframe(df.iloc[row])
                        results.append((f"Row {row}", df.iloc[row].copy()))

                elif axis == "Multiple Columns":
                    selected_columns = st.multiselect("Select columns to process", df.columns)
                    value = st.number_input("Enter the value:", value=1.0, key="multi_value_input")
                    if st.button("Apply Operation to Selected Columns"):
                        for column in selected_columns:
                            df[column] = perform_operation(df[column], operation_type, value)
                            st.write(f"Result for column {column}:")
                            st.dataframe(df[column])
                            results.append((f"Column {column}", df[column].copy()))

                elif axis == "Multiple Rows":
                    selected_rows = st.multiselect("Select rows to process", range(len(df)))
                    value = st.number_input("Enter the value:", value=1.0, key="multi_value_input_rows")
                    if st.button("Apply Operation to Selected Rows"):
                        for row in selected_rows:
                            df.iloc[row] = perform_operation(df.iloc[row], operation_type, value)
                            st.write(f"Result for row {row}:")
                            st.dataframe(df.iloc[row])
                            results.append((f"Row {row}", df.iloc[row].copy()))

            elif operation_type == "Sum":
                if axis == "Column":
                    column = st.selectbox("Select a column to sum", df.columns, key="sum_col")
                    if st.button("Apply Sum Operation to Column"):
                        total = df[column].sum()
                        st.write(f"Sum of column {column}: {total}")
                        results.append((f"Sum of Column {column}", total))

                elif axis == "Row":
                    row = st.number_input("Enter the row index to sum:", min_value=0, max_value=len(df)-1, step=1, key="sum_row")
                    if st.button("Apply Sum Operation to Row"):
                        total = df.iloc[row].sum()
                        st.write(f"Sum of row {row}: {total}")
                        results.append((f"Sum of Row {row}", total))

            elif operation_type == "Average":
                if axis == "Column":
                    column = st.selectbox("Select a column to average", df.columns, key="avg_col")
                    if st.button("Apply Average Operation to Column"):
                        avg = df[column].mean()
                        st.write(f"Average of column {column}: {avg}")
                        results.append((f"Average of Column {column}", avg))

                elif axis == "Row":
                    row = st.number_input("Enter the row index to average:", min_value=0, max_value=len(df)-1, step=1, key="avg_row")
                    if st.button("Apply Average Operation to Row"):
                        avg = df.iloc[row].mean()
                        st.write(f"Average of row {row}: {avg}")
                        results.append((f"Average of Row {row}", avg))

        # Apply operations and display intermediate results
        if st.button("Apply All Operations"):
            for op in operations:
                try:
                    operation, ax, target, *values = op
                    # Applying operations based on type
                    if operation == "Multiply":
                        if ax == "Column":
                            df[target] *= values[0]
                            results.append((f"Column {target}", df[target].copy()))
                        elif ax == "Row":
                            df.iloc[target] *= values[0]
                            results.append((f"Row {target}", df.iloc[target].copy()))
                    elif operation == "Add":
                        if ax == "Column":
                            df[target] += values[0]
                            results.append((f"Column {target}", df[target].copy()))
                        elif ax == "Row":
                            df.iloc[target] += values[0]
                            results.append((f"Row {target}", df.iloc[target].copy()))
                    elif operation == "Subtract":
                        if ax == "Column":
                            df[target] -= values[0]
                            results.append((f"Column {target}", df[target].copy()))
                        elif ax == "Row":
                            df.iloc[target] -= values[0]
                            results.append((f"Row {target}", df.iloc[target].copy()))
                    elif operation == "Divide":
                        if ax == "Column":
                            df[target] /= values[0]
                            results.append((f"Column {target}", df[target].copy()))
                        elif ax == "Row":
                            df.iloc[target] /= values[0]
                            results.append((f"Row {target}", df.iloc[target].copy()))
                except Exception as e:
                    st.error(f"Error applying operation {op}: {e}")
            st.success("All operations applied successfully!")

            # Display intermediate results
            for label, result in results:
                st.write(f"Result after modifying {label}:")
                st.dataframe(result)

        # Download Section
        st.markdown("""
        <div class="download-section">
            <h2>Step 4: Download Processed File</h2>
        </div>
        """, unsafe_allow_html=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
            writer.save()

        st.download_button(
            label="Download Processed Excel File",
            data=output.getvalue(),
            file_name="processed_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error loading file: {e}")

