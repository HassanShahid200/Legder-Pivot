import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

st.set_page_config(layout="wide", page_title="September Financial Dashboard")
st.title("September 2025 vs 2026 Ledger Dashboard")

uploaded_file = st.file_uploader("Apni SAP Excel File yahan upload karein ('General Ledger' aur 'Acct class' sheets zaroori hain)", type=['xlsx'])

if uploaded_file is not None:
    @st.cache_data
    def load_data(file):
        gl_df = pd.read_excel(file, sheet_name='General Ledger')
        acct_df = pd.read_excel(file, sheet_name='Acct class')
        
        gl_df['Posting Date'] = pd.to_datetime(gl_df['Posting Date'])
        
        sep_df = gl_df[gl_df['Posting Date'].dt.month == 9].copy()
        sep_df['Year'] = sep_df['Posting Date'].dt.year.astype(str)
        
        df = pd.merge(sep_df, acct_df, left_on='Acct Code', right_on='Account code', how='left')
        
        df['Current Main Class'] = df['Current Main Class'].fillna('Unclassified')
        df['Current Sub Class'] = df['Current Sub Class'].fillna('Unclassified')
        df['Title'] = df['Title'].fillna(df['Acct Name'])
        
        return df

    try:
        df = load_data(uploaded_file)
        
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
        
        df['Main_Class_Sort'] = pd.Categorical(df['Current Main Class'], categories=main_class_order, ordered=True)
        
        pivot_df = pd.pivot_table(
            df, 
            values='Deb./Cred. (LC)', 
            index=['Main_Class_Sort', 'Current Main Class', 'Current Sub Class', 'Title'], 
            columns='Year', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot_df.columns = [str(c) for c in pivot_df.columns]
        
        if '2025' in pivot_df.columns:
            pivot_df.rename(columns={'2025': 'Sep 2025'}, inplace=True)
        else:
            pivot_df['Sep 2025'] = 0.0

        if '2026' in pivot_df.columns:
            pivot_df.rename(columns={'2026': 'Sep 2026'}, inplace=True)
        else:
            pivot_df['Sep 2026'] = 0.0

        pivot_df['% Change'] = (pivot_df['Sep 2026'] - pivot_df['Sep 2025']) / pivot_df['Sep 2025'].replace(0, pd.NA)
        pivot_df['% Change'] = pivot_df['% Change'].fillna(0)

        pivot_df = pivot_df.sort_values(by=['Main_Class_Sort', 'Current Sub Class', 'Title'])
        pivot_df = pivot_df.drop(columns=['Main_Class_Sort'])
        
        pivot_df['Current Main Class'] = pivot_df['Current Main Class'].astype(str)
        pivot_df['Current Sub Class'] = pivot_df['Current Sub Class'].astype(str)
        pivot_df['Title'] = pivot_df['Title'].astype(str)

        st.subheader("1. September Summary Pivot")
        st.write("Table ko expand karein aur kisi bhi account ke amount par click karein.")

        gb = GridOptionsBuilder.from_dataframe(pivot_df)
        gb.configure_column('Current Main Class', rowGroup=True, hide=True)
        gb.configure_column('Current Sub Class', rowGroup=True, hide=True)
        gb.configure_column('Title', rowGroup=False)
        
        # FIXED: aggFunc='sum' add kiya hai taake groups blank na rahein
        gb.configure_column('Sep 2026', type=["numericColumn"], valueFormatter="x.toLocaleString('en-US', {minimumFractionDigits: 2})", aggFunc='sum')
        gb.configure_column('Sep 2025', type=["numericColumn"], valueFormatter="x.toLocaleString('en-US', {minimumFractionDigits: 2})", aggFunc='sum')
        gb.configure_column('% Change', valueFormatter="x.toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2})")
        
        # FIXED: use_checkbox=False kar diya hai. Ab amount/row par direct click kaam karega.
        gb.configure_selection(selection_mode='single', use_checkbox=False)
        gridOptions = gb.build()

        grid_response = AgGrid(
            pivot_df,
            gridOptions=gridOptions,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            enable_enterprise_modules=True,
            height=450,
            theme='streamlit'
        )

        selected_rows = grid_response['selected_rows']
        st.markdown("---")

        if selected_rows is not None and len(selected_rows) > 0:
            if isinstance(selected_rows, pd.DataFrame):
                selected_row = selected_rows.iloc[0]
            else:
                selected_row = selected_rows[0]
                
            selected_title = selected_row.get('Title')
            
            # Agar group header par click hua toh Title None/NaN hoga
            if pd.notna(selected_title) and selected_title != "":
                st.subheader(f"2. Fact Ledger Detail: {selected_title}")
                
                selected_year = st.radio("👇 Konsay saal ka poora ledger dekhna chahte hain?", ["2026", "2025"], horizontal=True)
                
                detail_df = df[(df['Title'] == selected_title) & (df['Year'] == selected_year)]
                
                cols_to_show = ['Acct Code', 'Posting Date', 'Origin', 'Doc No', 'Remarks', 'Offset Acct', 'Offset Acct Name', 'Deb./Cred. (LC)']
                detail_df['Posting Date'] = detail_df['Posting Date'].dt.strftime('%Y-%m-%d')
                
                st.write(f"**Total Transactions (Sep {selected_year}):** {len(detail_df)}")
                st.dataframe(
                    detail_df[cols_to_show].sort_values(by='Posting Date'),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("💡 Aapne ek Category Header select kiya hai. Detail dekhne ke liye usko expand karein aur andar kisi account ('Title') par click karein.")
        else:
            st.info("👆 Click to Ledger Amount")

    except Exception as e:
        st.error(f"❌ Error processing the file. (Details: {e})")
else:
    st.info("📂 Please Upload your Excel File")
