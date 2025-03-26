

import streamlit as st
import pandas as pd
import numpy as np
import io
st.markdown("""
**VBA Code Developed by:** Late B. N. Raviprakash  
**Translated to Python by:** Manojkumar Patil  
**Description:** This Streamlit-based application generates Markov Chain equations from user-uploaded Excel or CSV files.  
""")
class MarkovCalculator:
    def __init__(self):
        self.mcdata = None
        self.tot_rows = 0
        self.tot_cols = 0

    def calculate_markov(self, df, start_row, end_row, start_col, end_col):
        # Adjust indices for 0-based indexing
        start_row -= 1
        end_row -= 1
        start_col -= 1
        end_col -= 1
        
        self.tot_rows = end_row - start_row + 1
        self.tot_cols = end_col - start_col + 1
        self.mcdata = df.iloc[start_row:end_row+1, start_col:end_col+1].values
        
        # Calculate Markov chain
        mcdata1 = self.mcdata.copy()
        
        # Normalize and round data
        for n in range(self.tot_cols):
            for i in range(self.tot_rows):
                for j in range(self.tot_cols - 1):
                    mcdata1[i, j] = round((self.mcdata[i, j] / self.mcdata[i, -1]) * 10000) / 10000

        # Generate model text
        s_output = ""
        s_xend = ""
        s_minu = ""
        s_minv = ""
        i_rownum = 1

        for n in range(self.tot_cols - 1):
            for i in range(self.tot_rows - 1):
                for j in range(self.tot_cols - 1):
                    s_output += f"{mcdata1[i, j]}*x{chr(97 + j)}{n+1}+"
                    if j == 4:
                        s_output += "\n"
                    if i == 0:
                        if j < self.tot_cols - 2:
                            s_xend += f"1*x{chr(97 + n)}{j+1}+"
                        else:
                            s_xend += f"1*x{chr(97 + n)}{j+1}=1;"
                
                s_output += f"u{i_rownum}-v{i_rownum}={mcdata1[i+1, n]};\n"
                s_minu += f"u{i_rownum}+"
                
                if n == self.tot_cols - 2 and i == self.tot_rows - 2:
                    s_minv += f"v{i_rownum};"
                else:
                    s_minv += f"v{i_rownum}+"
                
                i_rownum += 1

            s_output += "\n"
            s_minu += "\n"
            s_minv += "\n"
            s_xend += "\n"

        model_text = "/* Objective function */\n"
        model_text += "min :\n"
        model_text += s_minu
        model_text += s_minv + "\n\n"
        model_text += "/* Variable bounds */\n"
        model_text += s_output + "\n"
        model_text += s_xend
        
        return model_text

def main():
    st.title("MARKOV EQUATION GENERATOR")
    
    # File upload - add CSV support
    uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=['xlsx', 'xls', 'xlsm', 'csv'])
    
    if uploaded_file is not None:
        # Handle both Excel and CSV files
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        # Show dataframe preview
        st.subheader("Data Preview")
        st.dataframe(df.head())
        
        # Display data dimensions
        st.info(f"📊 Data Dimensions: {df.shape[0]} rows × {df.shape[1]} columns")
        
        # Warning about last column
        st.warning("⚠️ Important: The last column in your selection must be the row total!")
        
        # Display column labels guide
        st.info("📌 Column Labels in your data:")
        col_labels = {i+1: label for i, label in enumerate(df.columns)}
        st.write(col_labels)
        
        # Range selection guide
        st.info("📌 Enter ranges as they appear in Excel. First column is usually the year/time period.")
        
        # Range specification
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Row Range")
            start_row = st.number_input("Start Row", min_value=1, value=1)
            end_row = st.number_input("End Row", min_value=start_row, value=len(df))
            
        with col2:
            st.subheader("Column Range")
            st.warning("⚠️ Important: The first column is ignored as it usually contains the years; hence, columns start from 2.")
            start_col = st.number_input("Start Column", min_value=2, value=2)
            end_col = st.number_input("End Column (max = " + str(len(df.columns)) + ")", min_value=start_col, value=len(df.columns))
        
        # Replace the duplicate output section with a single, styled output
        if st.button("Calculate Markov Chain"):
            calculator = MarkovCalculator()
            try:
                model_text = calculator.calculate_markov(df, start_row, end_row, start_col, end_col)
                
                # Display results in a single, styled container
                with st.container():
                    st.subheader("🎯 Model Output")
                    st.code(model_text, language="lingo")  # Using lingo syntax highlighting
                    
                    st.success("""
                    🚀 Your Markov Chain model is ready! Here's what to do next:
                    
                    1. Click the 'Copy' button in the top-right corner of the code box
                    2. Open LPSolve IDE
                    3. Clear the editor (Ctrl+A, Delete)
                    4. Paste the model (Ctrl+V)
                    5. Hit Run (F5) and watch the optimization magic happen! ✨
                    
                    Get ready to see your optimal solution! 💫
                    """)
                
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
