import streamlit as st
import pandas as pd
import io
import graphviz
import plotly.express as px

if 'operationStr' not in st.session_state:
    st.session_state.operationStr = "计算流（最多10个操作）: 当前值"

if 'operations' not in st.session_state:
    st.session_state.operations = []

if 'results' not in st.session_state:
    st.session_state.results = []

if 'filter_click_button' not in st.session_state:
    st.session_state.filter_click_button = False

if 'selectedSheet' not in st.session_state:
    st.session_state.selectedSheet = False

def filter_click_button():
    st.session_state.filter_click_button = True



# Function to perform the operation on data
def perform_operation(data, operation, value):
    if operation == "✖️":
        return data * value
    elif operation == "➕":
        return data + value
    elif operation == "➖":
        return data - value
    elif operation == "➗":
        return data / value
    else:
        return data

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
    <h2>Step 1: 上传 Excel</h2>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("选择文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Sheet Selection Section
        sheet_names = pd.ExcelFile(uploaded_file).sheet_names
        sheet = st.selectbox("请选择Sheet", sheet_names, index=None)

        if sheet in sheet_names:
            st.session_state.selectedSheet = True
        else:
            st.session_state.selectedSheet = False
            st.session_state.filter_click_button = False


        if st.session_state.selectedSheet:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)

            # Data Display Section
            st.markdown("""
            <div class="preview-section">
                <h2>Step 2: 数据预览</h2>
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(df)


            # Filter Section
            st.markdown("""
            <div class="operation-section">
                <h2>Step 3: 数据过滤</h2>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("Filter Data ", expanded=True):

                column = st.selectbox("选择要处理的列", df.columns, key="col_select")

                start_row = st.number_input("选择起始行（默认为0）", min_value=0, max_value=len(df) - 1, step=1,
                                            value=0)

                end_row = st.number_input("选择终止行（默认为最后）", min_value=start_row+1, max_value=len(df) - 1, step=1,
                                            value=len(df) - 1)

                df = df.loc[start_row:end_row]

                df = df.dropna(subset=[column])

                # Filtering the column based on a value range
                minVal = st.number_input(f"设置下限（默认为最小值）", value=df[column].min())
                maxVal = st.number_input(f"设置上限（默认为最大值）:", value=df[column].max())

                if st.button(f"过滤 {column}"):
                    filter_click_button()
                if st.session_state.filter_click_button:
                    filtered_df = df[(df[column] >= minVal) & (df[column] <= maxVal)]
                    st.write(f"过滤后的数据：{column} (范围: {minVal} - {maxVal})")
                    st.dataframe(filtered_df)
                    st.success(f"有效原数据总数：{len(df[column])}， "
                               f"过滤后数据总数：{len(filtered_df[column])}， "
                               f"过滤数据占比：{len(filtered_df[column]) / len(df[column]) * 100 :.3f} %，"
                               f"最小值：{filtered_df[column].min()}，"
                               f"最大值：{filtered_df[column].max()}，"
                               f"平均值：{filtered_df[column].mean():.3f}，"
                               f"中位数：{filtered_df[column].median()}")
                    df = filtered_df  # Set filtered data as the new dataframe


            # Operations Section
            st.markdown("""
            <div class="operation-section">
                <h2>Step 4: 定义计算流</h2>
            </div>
            """, unsafe_allow_html=True)
            with st.expander("Operation Stream", expanded=True):

                st.success(f"已选择要计算的列：{column}")
                operation_type = st.selectbox("选择运算", ["➕", "➖", "✖️", "➗"])
                value = st.number_input("添加常数", value=1.0, key="value_input")

                col1, col2, *_ = st.columns(10)

                with col1:
                    if st.button("添加计算"):
                        st.session_state.operations.append((operation_type, column, value))
                        st.session_state.operationStr += f"  ==>   {operation_type} {value}"

                with col2:
                    if st.button("重置计算"):
                        st.session_state.operations = []
                        st.session_state.operationStr = "计算流（最多10个操作）: 当前值"
                        st.session_state.results = []
                st.success(st.session_state.operationStr)



            # Apply operations and display intermediate results
            st.markdown("""
                    <div class="operation-section">
                        <h2>Step 5: 最终结果</h2>
                    </div>
                    """, unsafe_allow_html=True)
            with st.expander("Results", expanded=True):

                if st.button("计算"):
                    st.session_state.results = []
                    # add original data
                    st.session_state.results.append(df[column].copy())
                    for op in st.session_state.operations:
                        try:
                            operation, target, *values = op
                            df[target] = perform_operation(df[target], operation, values)
                            st.session_state.results.append(((values,operation,target), df[target].copy()))


                        except Exception as e:
                            st.error(f"Error applying operation {op}: {e}")
                    st.success("计算完成！步骤如下")

                    # Display intermediate results

                    res_colums = st.columns(11)
                    for index in range(len(st.session_state.results)):
                        with res_colums[index]:
                            if index > 0:
                                st.write(f"第{index}步: {st.session_state.results[index][0][1]} {st.session_state.results[index][0][0]}")
                                st.dataframe(st.session_state.results[index][1])
                            else:
                                st.write(f"初始值")
                                st.dataframe(st.session_state.results[index])


                    final_results = st.session_state.results[-1][1]
                    st.write(f"结果统计信息")
                    st.success(
                        f"最小值：{final_results.min():.3f}，"
                        f"最大值：{final_results.max():.3f}，"
                        f"平均值：{final_results.mean():.3f}，"
                        f"中位数：{final_results.median():.3f}，"
                        f"总和：{final_results.sum():.3f}"
                    )

                    fig = px.scatter(final_results)
                    fig.update_layout(
                        xaxis_title="X: 行",  # X轴的标签
                        yaxis_title="Y: 值 ",  # Y轴的标签
                    )
                    st.plotly_chart(fig)


            # Download Section
            st.markdown("""
            <div class="download-section">
                <h2>Step 6（可选）: 下载处理后的表格</h2>
            </div>
            """, unsafe_allow_html=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name=sheet, index=False)
                writer._save()

            st.download_button(
                label="Download Processed Excel File",
                data=output.getvalue(),
                file_name="processed_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error loading file: {e}")


