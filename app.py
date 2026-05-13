import streamlit as st
import pandas as pd
import numpy as np
from google import genai
import io

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI Retail Planning Assistant", page_icon="🛍️", layout="wide")
st.title("🛍️ End-to-End Retail Planning Console")
st.markdown("Upload raw supplier data to automatically generate the planning deliverable, then chat with the AI to analyze the results.")

# --- INITIALIZE SESSION STATE ---
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None
if "messages" not in st.session_state:
    st.session_state.messages = []

# --- SIDEBAR: SETUP ---
with st.sidebar:
    st.header("1. Setup & Upload")
    api_key = st.text_input("Enter your Gemini API Key:", type="password")
    uploaded_file = st.file_uploader("Upload Raw Reference Data (.xlsx)", type=["xlsx"])
    
    if st.button("Clear Chat History"):
        st.session_state.messages = []

# --- SECTION 1: DATA PROCESSING ---
if uploaded_file:
    st.header("⚙️ Step 1: Data Processing Pipeline")
    
    # We use a button to trigger processing so it doesn't run repeatedly on every interaction
    if st.button("Process Raw Data into Deliverable") or st.session_state.processed_df is not None:
        
        # Only run the heavy pandas math if we haven't already processed it
        if st.session_state.processed_df is None:
            with st.spinner("Cleaning data, merging tabs, and calculating forecasts..."):
                try:
                    # 1. LOAD THE DATA
                    df_file = pd.ExcelFile(uploaded_file)
                    tab1 = pd.read_excel(df_file, "TAB 1", header=1)
                    tab2 = pd.read_excel(df_file, "Tab 2", header=2)
                    tab3 = pd.read_excel(df_file, "Tab 3", header=3)
                    tab4 = pd.read_excel(df_file, "Tab 4", header=3)

                    # 2. CLEAN & PREP
                    tab1 = tab1[~tab1['Brand'].astype(str).str.contains("Total", na=False)]
                    tab2 = tab2[~tab2['Brand'].astype(str).str.contains("Total", na=False)]

                    tab1['Brand_Clean'] = tab1['Brand'].astype(str).str.title().str.strip()
                    tab2['Brand_Clean'] = tab2['Brand'].astype(str).str.title().str.strip()
                    tab3['Brand_Clean'] = tab3['BRAND'].astype(str).str.title().str.strip()
                    tab4['Brand_Clean'] = tab4['BRAND'].astype(str).str.title().str.strip()

                    tab3_grouped = tab3.groupby('Brand_Clean', as_index=False)['$ SHIPMENTS 10/23/23'].sum()

                    # 3. MERGE
                    df = tab1.merge(tab2[['Brand_Clean', 'Sum of OH+OO']], on='Brand_Clean', how='left')
                    df = df.merge(tab3_grouped[['Brand_Clean', '$ SHIPMENTS 10/23/23']], on='Brand_Clean', how='left')
                    df = df.merge(tab4[['Brand_Clean', 'Total Shipment $ Q1 2024 sets']], on='Brand_Clean', how='left')

                    # Replace NaN with 0 for numeric columns specifically to avoid warnings
                    numeric_cols = df.select_dtypes(include=[np.number]).columns
                    df[numeric_cols] = df[numeric_cols].fillna(0)

                    # 4. CALCULATIONS
                    df['% Change in YTD Sales'] = (df['YTD SET SALES\n(through to 9/22/2023)'] - df['PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)']) / df['PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)']
                    df['Total Expected Sales'] = df['Q3 2022\nSET SALES'] + df['Q4 2022\nSET SALES'] + df['Q1 2023 \nSET SALES']
                    df['Total Available Inventory'] = df['Sum of OH+OO'] + df['$ SHIPMENTS 10/23/23'] + df['Total Shipment $ Q1 2024 sets']
                    df['Dollar Difference'] = df['Total Available Inventory'] - df['Total Expected Sales']
                    df['Inventory as % of Expected Sales'] = df['Total Available Inventory'] / df['Total Expected Sales']
                    df['Comments'] = ""

                    # 5. CLEAN UP
                    final_columns = [
                        'Axe', 'Brand', 
                        'YTD SET SALES\n(through to 9/22/2023)', 
                        'PRIOR YEAR YTD SET SALES\n(as of same time last year; through to 9/21/2022)',
                        '% Change in YTD Sales', 
                        'Q3 2022\nSET SALES', 
                        'Q4 2022\nSET SALES', 
                        'Q1 2023 \nSET SALES', 
                        'Total Expected Sales', 
                        'Sum of OH+OO', 
                        '$ SHIPMENTS 10/23/23', 
                        'Total Shipment $ Q1 2024 sets',
                        'Dollar Difference', 
                        'Inventory as % of Expected Sales', 
                        'Comments'
                    ]
                    
                    # Save to session state so we don't lose it
                    st.session_state.processed_df = df[final_columns]
                    st.success("Pipeline complete! Data is ready.")
                    
                except Exception as e:
                    st.error(f"An error occurred during processing: {e}")

        # If processing was successful, show the download button and preview
        if st.session_state.processed_df is not None:
            st.dataframe(st.session_state.processed_df.head(5))
            
            # Convert dataframe to an Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                st.session_state.processed_df.to_excel(writer, index=False)
            processed_data_bytes = output.getvalue()
            
            st.download_button(
                label="📥 Download Deliverable.xlsx",
                data=processed_data_bytes,
                file_name="Python_Deliverable.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.divider()

# --- SECTION 2: AI CHATBOT (Unlocks after processing) ---
if st.session_state.processed_df is not None:
    st.header("🤖 Step 2: AI Planning Insights")
    
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar to unlock the chatbot.")
    else:
        # Initialize client
        client = genai.Client(api_key=api_key)
        
        # Prepare the context string from the processed data
        # Using to_csv instead of markdown saves tokens for the Lite model
        data_string = st.session_state.processed_df.to_csv(index=False)
        
        system_instructions = f"""
        You are an expert Planning Manager for a national accounts team at a cosmetics brand. 
        Your goal is to analyze retailer sales, identify inventory risks, and suggest actionable strategies.
        
        Here is the finalized deliverable data you just generated:
        
        {data_string}
        
        Answer the user's questions based strictly on this data. Look at Dollar Difference and 
        Inventory as % of Expected Sales to find risks. Be concise, professional, and strategic.
        """

        # Display chat history
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Accept user input
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
elif not uploaded_file:
    st.info("👈 Please upload the raw reference data in the sidebar to begin.")