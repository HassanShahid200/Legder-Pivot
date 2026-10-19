import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(layout="wide", page_title="September Financial Dashboard")
st.title("September 2025 vs 2026 Ledger Dashboard")

# 1. Add the Drag-and-Drop File Uploader
uploaded_file = st.file_uploader("Upload your SAP Excel File (must contain 'General Ledger' and 'Acct class' sheets)", type=['xlsx'])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        # Load data directly from the uploaded file in memory
        gl_df = pd.read_excel(file, sheet_name='General Ledger')
        acct_df = pd.read_excel(file, sheet_name='Acct class')
        
        # Parse dates
        gl_df['Posting Date'] = pd.to_datetime(gl_df['Posting Date'])
        
        # Filter exclusively for September (Month == 9)
        sep_df = gl_df[gl_df['Posting Date'].dt.month == 9].copy()
        sep_df['Year'] = sep_df['Posting Date'].dt.year.astype(str)
        
        # Merge with Account Classes
        df = pd.merge(sep_df, acct_df, left_on='Acct Code', right_on='Account code', how='left')
        
        df['Current Main Class'] = df['Current Main Class'].fillna('Unclassified')
        df['Current Sub Class'] = df['Current Sub Class'].fillna('Unclassified')
        df['Title'] = df['Title'].fillna(df['Acct Name'])
        
        # Enforce exact sorting sequence
        main_class_order = [
            'Sales - net',
            'Cost of sales',
            'Selling and distribution expenses',
            'Administrative expenses',
            'Other operating incomes',
            'Other operating expenses',
            'Finance Costs',
            'Taxation'
        ]
        
        df['Current Main Class'] = pd.Categorical(df['Current Main Class'], categories=main_class_order, ordered=True)
        return df

    try:
        # Load the dataframe using the uploaded file
        df = load_data(uploaded_file)
        
        # Pivot the September Data
        pivot_df = pd.pivot_table(
            df, 
            values='Deb./Cred. (LC)', 
            index=['Current Main Class', 'Current Sub Class', 'Title'], 
            columns='Year', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # Rename year columns
        if '2025' in pivot_df.columns:
            pivot_df.rename(columns={'2025': 'Sep 2025'}, inplace=True)
        else:
            pivot_df['Sep 2025'] = 0.0

        if '2026' in pivot_df.columns:
            pivot_df.rename(columns={'2026': 'Sep 2026'}, inplace=True)
        else:
            pivot_df['Sep 2026'] = 0.0

        # Calculate % Change
        pivot_df['% Change'] = (pivot_df['Sep 2026'] - pivot_df['Sep 2025']) / pivot_df['Sep 2025'].replace(0, pd.NA)
        pivot_df['% Change'] = pivot_df['% Change'].fillna(0)

        # Sort by the locked categorical order
        pivot_df = pivot_df.sort_values(by=['Current Main Class', 'Current Sub Class', 'Title'])

        st.subheader("1. September Summary Pivot")
        st.write("Expand the classes. Select an account row to view its Sep 2026 ledger below.")

        # Configure Ag-Grid
        gb = GridOptionsBuilder.from_dataframe(pivot_df)
        gb.configure_column('Current Main Class', rowGroup=True, hide=True)
        gb.configure_column('Current Sub Class', rowGroup=True, hide=True)
        gb.configure_column('Title', rowGroup=False)
        gb.configure_column('Sep 2026', type=["numericColumn"], valueFormatter="x.toLocaleString('en-US', {minimumFractionDigits: 2})")
        gb.configure_column('Sep 2025', type=["numericColumn"], valueFormatter="x.toLocaleString('en-US', {minimumFractionDigits: 2})")
        gb.configure_column('% Change', valueFormatter="x.toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2})")
        gb.configure_selection(selection_mode='single', use_checkbox=True)
        
        gridOptions = gb.build()

        grid_response = AgGrid(
            pivot_df,
            gridOptions=gridOptions,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            enable_enterprise_modules=True,
            height=450,
            theme='streamlit'
        )

        # Display Full Ledger Fact Table for Sep 2026
        selected_rows = grid_response['selected_rows']
        st.markdown("---")

        if selected_rows is not None and len(selected_rows) > 0:
            if isinstance(selected_rows, pd.DataFrame):
                selected_row = selected_rows.iloc[0]
            else:
                selected_row = selected_rows[0]
                
            selected_title = selected_row.get('Title')
            
            if selected_title:
                st.subheader(f"2. Fact Ledger for: {selected_title} (Sep 2026)")
                detail_df = df[(df['Title'] == selected_title) & (df['Year'] == '2026')]
                cols_to_show = ['Acct Code', 'Posting Date', 'Origin', 'Doc No', 'Remarks', 'Offset Acct', 'Offset Acct Name', 'Deb./Cred. (LC)']
                detail_df['Posting Date'] = detail_df['Posting Date'].dt.strftime('%Y-%m-%d')
                
                st.dataframe(
                    detail_df[cols_to_show].sort_values(by='Posting Date'),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("💡 You selected a group header. Please expand the group and select a specific 'Title' row.")
        else:
            st.info("👆 Select an account in the table above to view its detailed Sep 2026 transactions.")

    except ValueError as e:
        st.error(f"❌ Error processing the file. Please ensure it contains both 'General Ledger' and 'Acct class' sheets. (Details: {e})")

else:
    # Instructions to display before a file is uploaded
    st.info("📂 Please drag and drop your exported SAP Excel file into the box above to generate the dashboard.")
