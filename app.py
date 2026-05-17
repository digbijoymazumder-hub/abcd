import streamlit as st
import pandas as pd
import numpy as np
from google import genai
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI Retail Planning Assistant", page_icon="🛍️", layout="wide")
st.title("🛍️ End-to-End Retail Planning Console")
st.markdown("\n Groupwork Deployment By \n 1. Puspanjali Dahal \n 2. Harishraghavend Balaji \n 3. Dhruvi Malesha \n 4. Ming Fang \n 5. Moosa Ali \n 6. Digbijoy Mazumder \n \n Upload raw supplier data to automatically generate the planning deliverable, then chat with the AI to analyze the results.")

# --- INITIALIZE SESSION STATE ---
if "raw_file_bytes" not in st.session_state:
    st.session_state.raw_file_bytes = None
if "raw_file_name" not in st.session_state:
    st.session_state.raw_file_name = None
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None
if "messages" not in st.session_state:
    st.session_state.messages = []

# --- SIDEBAR: SETUP ---
with st.sidebar:
    st.header("1. Setup & Upload")
    
    if st.session_state.raw_file_bytes is None:
        uploaded_file = st.file_uploader("Upload Raw Reference Data (.xlsx)", type=["xlsx"])
        
        if uploaded_file is not None:
            st.session_state.raw_file_bytes = uploaded_file.getvalue()
            st.session_state.raw_file_name = uploaded_file.name
            st.rerun()
    else:
        st.success(f"✅ File uploaded: {st.session_state.raw_file_name}")
        st.caption("Further uploads disabled to protect current session.")
        
        if st.button("Start Over (Reset App)", type="primary"):
            st.session_state.clear()
            st.rerun()
            
    st.divider()
    if st.button("Clear Chat History"):
        st.session_state.messages = []

# --- SECTION 1: DATA PROCESSING ---
if st.session_state.raw_file_bytes is not None:
    st.header("⚙️ Step 1: Data Processing Pipeline")
    
    if st.button("Process Raw Data into Deliverable") or st.session_state.processed_df is not None:
        
        if st.session_state.processed_df is None:
            with st.spinner("Cleaning data, calculating forecasts, and formatting Excel formulas..."):
                try:
                    # 1. LOAD THE DATA
                    df_file = pd.ExcelFile(io.BytesIO(st.session_state.raw_file_bytes))
                    tab1 = pd.read_excel(df_file, "TAB 1", header=1)
                    tab2 = pd.read_excel(df_file, "Tab 2", header=2)
                    tab3 = pd.read_excel(df_file, "Tab 3", header=3)
                    tab4 = pd.read_excel(df_file, "Tab 4", header=3)

                    # 2. CLEAN & PREP (CRITICAL: Remove all garbage blank rows first)
                    tab1 = tab1.dropna(subset=['Brand'])
                    tab1 = tab1[~tab1['Brand'].astype(str).str.contains("Total", case=False, na=False)]
                    tab2 = tab2[~tab2['Brand'].astype(str).str.contains("Total", case=False, na=False)]

                    # Forward fill the Axe category to all valid brands
                    tab1['Axe'] = tab1['Axe'].ffill()

                    tab1['Brand_Clean'] = tab1['Brand'].astype(str).str.title().str.strip()
                    tab2['Brand_Clean'] = tab2['Brand'].astype(str).str.title().str.strip()
                    tab3['Brand_Clean'] = tab3['BRAND'].astype(str).str.title().str.strip()
                    tab4['Brand_Clean'] = tab4['BRAND'].astype(str).str.title().str.strip()

                    tab3_grouped = tab3.groupby('Brand_Clean', as_index=False)['$ SHIPMENTS 10/23/23'].sum()

                    # 3. MERGE
                    df = tab1.merge(tab2[['Brand_Clean', 'Sum of OH+OO']], on='Brand_Clean', how='left')
                    df = df.merge(tab3_grouped[['Brand_Clean', '$ SHIPMENTS 10/23/23']], on='Brand_Clean', how='left')
                    df = df.merge(tab4[['Brand_Clean', 'Total Shipment $ Q1 2024 sets']], on='Brand_Clean', how='left')

                    numeric_cols = df.select_dtypes(include=[np.number]).columns
                    df[numeric_cols] = df[numeric_cols].fillna(0)

                    # 4. BRAND LEVEL CALCULATIONS
                    df['% Change in YTD Sales'] = np.where(
                        df['PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)'] != 0,
                        (df['YTD SET SALES\n(through to 9/22/2023)'] - df['PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)']) / df['PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)'],
                        np.nan
                    )
                    df['Total Expected Sales'] = df['Q3 2022\nSET SALES'] + df['Q4 2022\nSET SALES'] + df['Q1 2023 \nSET SALES']
                    df['Total Available Inventory'] = df['Sum of OH+OO'] + df['$ SHIPMENTS 10/23/23'] + df['Total Shipment $ Q1 2024 sets']
                    df['Dollar Difference'] = df['Total Available Inventory'] - df['Total Expected Sales']
                    
                    df['Inventory as % of Expected Sales'] = np.where(
                        df['Total Expected Sales'] != 0,
                        df['Total Available Inventory'] / df['Total Expected Sales'],
                        np.nan
                    )
                    df['Comments'] = ""

                    # 5. DYNAMIC EXCEL FORMULA INJECTION
                    # Grab unique axes dynamically so we don't skip anything due to spelling differences
                    axes = df['Axe'].dropna().unique()
                    ordered_dfs = [] 
                    
                    # Excel rows: 1=Title, 2=Headers, 3=First Data Row
                    current_row = 3  
                    subtotal_rows = []

                    for axe in axes:
                        axe_df = df[df['Axe'] == axe].copy()
                        if axe_df.empty: continue
                        
                        num_brands = len(axe_df)
                        start_row = current_row
                        end_row = current_row + num_brands - 1
                        total_row_idx = end_row + 1
                        
                        # Wipe axis name for all but the first row of the block to match reference
                        axe_df.loc[axe_df.index[1:], 'Axe'] = ''
                        
                        # Build Native Excel formulas for the subtotal
                        sum_series = {
                            'Axe': f'{axe} Total',
                            'Brand': '',
                            'YTD SET SALES\n(through to 9/22/2023)': f'=SUM(C{start_row}:C{end_row})',
                            'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)': f'=SUM(D{start_row}:D{end_row})',
                            '% Change in YTD Sales': f'=IF(D{total_row_idx}=0, "", (C{total_row_idx}-D{total_row_idx})/D{total_row_idx})',
                            'Q1 2023 \nSET SALES': f'=SUM(F{start_row}:F{end_row})',
                            'Q4 2022\nSET SALES': f'=SUM(G{start_row}:G{end_row})',
                            'Q3 2022\nSET SALES': f'=SUM(H{start_row}:H{end_row})',
                            'Total Expected Sales': f'=SUM(I{start_row}:I{end_row})',
                            'Sum of OH+OO': f'=SUM(J{start_row}:J{end_row})',
                            '$ SHIPMENTS 10/23/23': f'=SUM(K{start_row}:K{end_row})',
                            'Total Shipment $ Q1 2024 sets': f'=SUM(L{start_row}:L{end_row})',
                            'Dollar Difference': f'=J{total_row_idx}+K{total_row_idx}+L{total_row_idx}-I{total_row_idx}',
                            'Inventory as % of Expected Sales': f'=IF(I{total_row_idx}=0, "", (J{total_row_idx}+K{total_row_idx}+L{total_row_idx})/I{total_row_idx})',
                            'Comments': ''
                        }
                        
                        ordered_dfs.append(axe_df)
                        ordered_dfs.append(pd.DataFrame([sum_series]))
                        
                        subtotal_rows.append(total_row_idx)
                        current_row = total_row_idx + 1

                    # Calculate Grand Total with dynamic formula linking
                    grand_total = {
                        'Axe': 'Grand Total',
                        'Brand': '',
                        'YTD SET SALES\n(through to 9/22/2023)': f'={"+".join([f"C{r}" for r in subtotal_rows])}',
                        'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)': f'={"+".join([f"D{r}" for r in subtotal_rows])}',
                        '% Change in YTD Sales': f'=IF(D{current_row}=0, "", (C{current_row}-D{current_row})/D{current_row})',
                        'Q1 2023 \nSET SALES': f'={"+".join([f"F{r}" for r in subtotal_rows])}',
                        'Q4 2022\nSET SALES': f'={"+".join([f"G{r}" for r in subtotal_rows])}',
                        'Q3 2022\nSET SALES': f'={"+".join([f"H{r}" for r in subtotal_rows])}',
                        'Total Expected Sales': f'={"+".join([f"I{r}" for r in subtotal_rows])}',
                        'Sum of OH+OO': f'={"+".join([f"J{r}" for r in subtotal_rows])}',
                        '$ SHIPMENTS 10/23/23': f'={"+".join([f"K{r}" for r in subtotal_rows])}',
                        'Total Shipment $ Q1 2024 sets': f'={"+".join([f"L{r}" for r in subtotal_rows])}',
                        'Dollar Difference': f'=J{current_row}+K{current_row}+L{current_row}-I{current_row}',
                        'Inventory as % of Expected Sales': f'=IF(I{current_row}=0, "", (J{current_row}+K{current_row}+L{current_row})/I{current_row})',
                        'Comments': ''
                    }
                    ordered_dfs.append(pd.DataFrame([grand_total]))

                    # Stitch all blocks together
                    df_with_totals = pd.concat(ordered_dfs, ignore_index=True)

                    # 6. EXACT REFERENCE FILE RENAMING
                    rename_cols = {
                        'Axe': 'Axis',
                        'Brand': 'Brand',
                        'YTD SET SALES\n(through to 9/22/2023)': 'YTD SET SALES\n(through to 9/22/2023)',
                        'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)': 'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)',
                        '% Change in YTD Sales': '% CHANGE IN YTD SET SALES',
                        'Q1 2023 \nSET SALES': 'Q1 2023 \nSET SALES',
                        'Q4 2022\nSET SALES': 'Q4 2022\nSET SALES',
                        'Q3 2022\nSET SALES': 'Q3 2022\nSET SALES',
                        'Total Expected Sales': "TOTAL EXPECTED SALES \n(through to end of Q1'2023)",
                        'Sum of OH+OO': "OH + OO ($)\n(as of 9/22/2023)",
                        '$ SHIPMENTS 10/23/23': "EXPECTED SHIPMENTS\n(Oct. 2023)",
                        'Total Shipment $ Q1 2024 sets': "EXPECTED SHIPMENTS\n (Q1 2023)",
                        'Dollar Difference': "TOTAL OH + OO ($) vs. TOTAL EXPECTED SALES",
                        'Inventory as % of Expected Sales': "TOTAL OH + OO ($)\n as % of TOTAL EXPECTED SALES",
                        'Comments': "COMMENTS"
                    }
                    
                    df_with_totals = df_with_totals.rename(columns=rename_cols)
                    
                    final_ordered_columns = [
                        'Axis', 'Brand', 
                        'YTD SET SALES\n(through to 9/22/2023)', 
                        'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)',
                        '% CHANGE IN YTD SET SALES', 
                        'Q1 2023 \nSET SALES', 
                        'Q4 2022\nSET SALES', 
                        'Q3 2022\nSET SALES', 
                        "TOTAL EXPECTED SALES \n(through to end of Q1'2023)", 
                        "OH + OO ($)\n(as of 9/22/2023)", 
                        "EXPECTED SHIPMENTS\n(Oct. 2023)", 
                        "EXPECTED SHIPMENTS\n (Q1 2023)",
                        "TOTAL OH + OO ($) vs. TOTAL EXPECTED SALES", 
                        "TOTAL OH + OO ($)\n as % of TOTAL EXPECTED SALES", 
                        "COMMENTS"
                    ]
                    
                    st.session_state.processed_df = df_with_totals[final_ordered_columns]
                    st.success("Pipeline complete! Data is ready.")
                    
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")

        # --- EXCEL FORMATTING (XLSXWRITER) ---
        if st.session_state.processed_df is not None:
            # We display the dataframe in Streamlit (it will show the raw formula strings)
            st.dataframe(st.session_state.processed_df)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                st.session_state.processed_df.to_excel(writer, sheet_name='Set Sales 2023', index=False, startrow=1)
                workbook = writer.book
                worksheet = writer.sheets['Set Sales 2023']
                
                # Formats
                title_format = workbook.add_format({'bold': True, 'font_size': 12})
                money_format = workbook.add_format({'num_format': '$#,##0.0', 'border': 1})
                percent_format = workbook.add_format({'num_format': '0.0%', 'border': 1})
                general_border = workbook.add_format({'border': 1})
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
                red_font = workbook.add_format({'font_color': 'red', 'num_format': '$#,##0.0', 'border': 1})

                worksheet.write('A1', 'Beutist SET SELLING 2023', title_format)

                for col_num, value in enumerate(st.session_state.processed_df.columns):
                    worksheet.write(1, col_num, value, header_format)
                    
                money_cols = [2, 3, 5, 6, 7, 8, 9, 10, 11, 12] 
                percent_cols = [4, 13] 
                
                for i in range(len(st.session_state.processed_df.columns)):
                    if i in money_cols:
                        worksheet.set_column(i, i, 16, money_format)
                    elif i in percent_cols:
                        worksheet.set_column(i, i, 14, percent_format)
                    else:
                        worksheet.set_column(i, i, 15, general_border)

                worksheet.conditional_format(2, 12, len(st.session_state.processed_df) + 1, 12, {
                    'type': 'cell',
                    'criteria': '<',
                    'value': 0,
                    'format': red_font
                })

            processed_data_bytes = output.getvalue()
            
            st.download_button(
                label="📥 Download Formatted Deliverable.xlsx",
                data=processed_data_bytes,
                file_name="Python_Deliverable_Formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.divider()

# --- SECTION 2: AI CHATBOT ---
if st.session_state.processed_df is not None:
    st.header("🤖 Step 2: AI Planning Insights")
    
    api_key = st.text_input("🔑 Enter your Gemini API Key to unlock the Chatbot:", type="password")
    
    if not api_key:
        st.info("Please enter your API key above to start analyzing your generated report.")
    else:
        client = genai.Client(api_key=api_key)
        
        data_string = st.session_state.processed_df.to_csv(index=False)
        
        system_instructions = f"""
        You are an expert Planning Manager for a national accounts team at a cosmetics brand. 
        Your goal is to analyze retailer sales, identify inventory risks, and suggest actionable strategies.
        
        Here is the finalized deliverable data you just generated:
        
        {data_string}
        
        Answer the user's questions based strictly on this data. Look at "TOTAL OH + OO ($) vs. TOTAL EXPECTED SALES" and 
        "TOTAL OH + OO ($) as % of TOTAL EXPECTED SALES" to find risks. Be concise, professional, and strategic.
        """

        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        if prompt := st.chat_input("Ask a question about the generated inventory report..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                try:
                    response = client.models.generate_content(
                        model="gemini-3.1-flash-lite", 
                        contents=f"{system_instructions}\n\nUser Question: {prompt}"
                    )
                    ai_reply = response.text
                    message_placeholder.markdown(ai_reply)
                    st.session_state.messages.append({"role": "assistant", "content": ai_reply})
                except Exception as e:
                    st.error(f"An error occurred with the AI: {e}")
elif st.session_state.raw_file_bytes is None:
    st.info("👈 Please upload the raw reference data in the sidebar to begin.")
