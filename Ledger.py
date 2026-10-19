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
        
        # Sirf September (Month == 9) ka data filter karein
        sep_df = gl_df[gl_df['Posting Date'].dt.month == 9].copy()
        sep_df['Year'] = sep_df['Posting Date'].dt.year.astype(str)
        
        df = pd.merge(sep_df, acct_df, left_on='Acct Code', right_on='Account code', how='left')
        
        df['Current Main Class'] = df['Current Main Class'].fillna('Unclassified')
        df['Current Sub Class'] = df['Current Sub Class'].fillna('Unclassified')
        df['Title'] = df['Title'].fillna(df['Acct Name'])
        
        return df

    try:
        df = load_data(uploaded_file)
        
        # Sequence lock karne ke liye Categorical sorting
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
        
        # Pivot table banayein
        pivot_df = pd.pivot_table(
            df, 
            values='Deb./Cred. (LC)', 
            index=['Main_Class_Sort', 'Current Main Class', 'Current Sub Class', 'Title'], 
            columns='Year', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        # Columns ko explicitly text mein convert karein taake blank screen issue na aaye
        pivot_df.columns = [str(c) for c in pivot_df.columns]
        
        if '2025' in pivot_df.columns:
            pivot_df.rename(columns={'2025': 'Sep 2025'}, inplace=True)
        else:
            pivot_df['Sep 2025'] = 0.0

        if '2026' in pivot_df.columns:
            pivot_df.rename(columns={'2026': 'Sep 2026'}, inplace=True)
        else:
            pivot_df['Sep 2026'] = 0.0

        # % Change Formula
        pivot_df['% Change'] = (pivot_df['Sep 2026'] - pivot_df['Sep 2025']) / pivot_df['Sep 2025'].replace(0, pd.NA)
        pivot_df['% Change'] = pivot_df['% Change'].fillna(0)

        # Sort karne ke baad 'Main_Class_Sort' drop kar dein taake AgGrid crash na ho
        pivot_df = pivot_df.sort_values(by=['Main_Class_Sort', 'Current Sub Class', 'Title'])
        pivot_df = pivot_df.drop(columns=['Main_Class_Sort'])
        
        # Tamam classes ko strictly string banayein
        pivot_df['Current Main Class'] = pivot_df['Current Main Class'].astype(str)
        pivot_df['Current Sub Class'] = pivot_df['Current Sub Class'].astype(str)
        pivot_df['Title'] = pivot_df['Title'].astype(str)

        st.subheader("1. September Summary Pivot")
        st.write("Table ko expand karein aur account row ko tick karein.")

        # Interactive AgGrid Setup
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

        selected_rows = grid_response['selected_rows']
        st.markdown("---")

        if selected_rows is not None and len(selected_rows) > 0:
            if isinstance(selected_rows, pd.DataFrame):
                selected_row = selected_rows.iloc[0]
            else:
                selected_row = selected_rows[0]
                
            selected_title = selected_row.get('Title')
            
            if selected_title:
                st.subheader(f"2. Fact Ledger Detail: {selected_title}")
                
                # Naya Feature: Saal choose karne ke liye button!
                selected_year = st.radio("👇 Konsay saal ka poora ledger dekhna chahte hain?", ["2026", "2025"], horizontal=True)
                
                # Sirf selected saal aur selected account ka data filter hoga
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
                st.info("💡 Aapne ek Category select ki hai. Poora ledger dekhne ke liye plus (+) se usko expand karein aur andar se kisi ek 'Title' ko select karein.")
        else:
            st.info("👆 Upar wali table se account select karein taake uska Fact Ledger yahan show ho sake.")

    except Exception as e:
        st.error(f"❌ Error processing the file. (Details: {e})")
else:
    st.info("📂 Upload button par click karein aur file attach karein.")
