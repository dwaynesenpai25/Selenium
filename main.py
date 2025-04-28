import pandas as pd
import streamlit as st
from tabs.dar import Dar  # Assuming you have a Dar class in dar.py
from tabs.skips import Skips  # Assuming you have a Skips class in skips.py

def main():
    st.title("DAR & Skips Automation")
    tabs = st.tabs(["DAR", "SKIPS"])

    with tabs[0]:  # DAR Tab
        st.header("DAR Automation")
        uploaded_file_dar = st.file_uploader("Upload Excel file for DAR", type=["xlsx"], key="dar_uploader")

        if uploaded_file_dar is not None:
            st.success("File uploaded successfully for DAR. Please click 'Start Processing' to begin.")
            if st.button("Start Processing", key="dar_start"):
                try:
                    df_dar = pd.read_excel(uploaded_file_dar)
                    df_dar = df_dar.astype(str)

                    st.write("DAR Automation started...")
                    start = Dar()
                    start.main(df_dar)

                except Exception as e:
                    st.error(f"Error processing DAR file: {str(e)}")
        else:
            st.error("Please upload an Excel file for DAR to proceed.")

    with tabs[1]:  # Skips Tab
        st.header("Skips Automation")
        uploaded_file_skips = st.file_uploader("Upload Excel file for Skips", type=["xlsx"], key="skips_uploader")

        if uploaded_file_skips is not None:
            st.success("File uploaded successfully for Skips. Please click 'Start Processing' to begin.")
            if st.button("Start Processing", key="skips_start"):
                try:
                    df_skips = pd.read_excel(uploaded_file_skips)
                    df_skips = df_skips.astype(str)
                    st.write("Dataframe for Skips:", df_skips)

                    st.write("Skips Automation started...")
                    start = Skips()
                    start.main(df_skips)

                    # Placeholder for Skips logic - you would need to define what "Skips" automation does
                    st.warning("Skips automation is not implemented yet. Please add your logic here.")
                    # If you have a different automation process for Skips, replace the warning with that logic

                except Exception as e:
                    st.error(f"Error processing Skips file: {str(e)}")
        else:
            st.error("Please upload an Excel file for Skips to proceed.")

if __name__ == '__main__':
    main()