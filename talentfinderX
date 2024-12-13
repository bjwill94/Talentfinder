import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


def calculate_total_and_rank(df, point_values):
    """
    Calculate total marks, ranks and points using configurable point values
    """
    # Calculate total marks
    df['total marks'] = df['mark1'].fillna(0) + df['mark2'].fillna(0) + df['mark3'].fillna(0)

    # Sort by total marks in descending order
    df = df.sort_values(by='total marks', ascending=False)

    # Calculate ranks (higher marks get better rank)
    df['Rank'] = df['total marks'].rank(method='min', ascending=False)

    # Add points based on rank using configured point values
    df['Points'] = 0  # Initialize all points to 0
    df.loc[df['Rank'] == 1, 'Points'] = point_values['first']
    df.loc[df['Rank'] == 2, 'Points'] = point_values['second']
    df.loc[df['Rank'] == 3, 'Points'] = point_values['third']

    # Check for ties
    tie_groups = identify_ties(df)

    # Add tie indicator column
    df['Tie Status'] = ''
    for mark, chest_numbers in tie_groups.items():
        df.loc[df['Chest No'].isin(chest_numbers), 'Tie Status'] = 'Tied'

    return df, tie_groups


def identify_ties(df):
    """
    Identify ties in total marks and return tied entries, excluding entries with 0 marks
    """
    # Filter out entries with 0 total marks
    df_non_zero = df[df['total marks'] > 0]

    # Group by total marks and find groups with more than one entry
    ties = df_non_zero[df_non_zero['total marks'].duplicated(keep=False)].sort_values('total marks', ascending=False)

    # Create a dictionary of tied marks and corresponding chest numbers
    tie_groups = {}
    if not ties.empty:
        for mark in ties['total marks'].unique():
            tied_entries = ties[ties['total marks'] == mark]
            tie_groups[mark] = list(tied_entries['Chest No'])

    return tie_groups


def merge_with_master_data(marks_df, master_df):
    # Ensure column names match for merging
    if 'Chest No' in marks_df.columns and 'Chest No' in master_df.columns:
        # Merge the dataframes based on Chest No
        merged_df = pd.merge(
            master_df,
            marks_df[['Chest No', 'mark1', 'mark2', 'mark3', 'total marks', 'Rank', 'Points', 'Tie Status']],
            on='Chest No',
            how='right'
        )
        # Sort by total marks in descending order
        merged_df = merged_df.sort_values(by='total marks', ascending=False)
        return merged_df
    else:
        st.error("Chest No column not found in one or both files")
        return None


def display_tie_information(tie_groups):
    if tie_groups:
        st.warning("⚠️ Ties Detected!")
        st.write("The following participants have tied scores:")
        for mark, chest_numbers in tie_groups.items():
            st.write(f"Total Mark: {mark}")
            st.write(f"Tied Chest Numbers: {', '.join(map(str, chest_numbers))}")
            st.write("---")
    else:
        st.success("No ties detected in the results (excluding entries with 0 marks).")


def main():
    st.title("Event Results Calculator")

    # Get event name
    event_name = st.text_input("Enter Event Name", "")

    # Points Configuration Section
    st.subheader("Points Configuration")
    col1, col2, col3 = st.columns(3)

    with col1:
        first_place_points = st.number_input("Points for 1st Place", min_value=0, value=5, step=1)
    with col2:
        second_place_points = st.number_input("Points for 2nd Place", min_value=0, value=3, step=1)
    with col3:
        third_place_points = st.number_input("Points for 3rd Place", min_value=0, value=1, step=1)

    point_values = {
        'first': first_place_points,
        'second': second_place_points,
        'third': third_place_points
    }

    # File Upload Section
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Upload Marks Entry File")
        marks_file = st.file_uploader("Upload Excel file with Chest Numbers", type=['xlsx', 'xls'])

    with col2:
        st.subheader("Upload Master Data")
        master_file = st.file_uploader("Upload Excel file with participant details", type=['xlsx', 'xls'])

    if marks_file is not None:
        try:
            marks_df = pd.read_excel(marks_file)

            # Verify required columns exist
            required_columns = ['Chest No', 'mark1', 'mark2', 'mark3', 'total marks', 'Rank', 'Points', 'Tie Status']
            missing_columns = [col for col in required_columns if col not in marks_df.columns]

            # Add missing columns if necessary
            for col in missing_columns:
                marks_df[col] = np.nan

            # Show the marks dataframe as an editable table
            st.write("Enter marks in the table below:")
            edited_df = st.data_editor(
                marks_df,
                num_rows="dynamic",
                disabled=["Chest No", "total marks", "Rank", "Points", "Tie Status"],
                column_config={
                    "mark1": st.column_config.NumberColumn(
                        "Mark 1",
                        min_value=0,
                        max_value=100,
                        step=0.5,
                    ),
                    "mark2": st.column_config.NumberColumn(
                        "Mark 2",
                        min_value=0,
                        max_value=100,
                        step=0.5,
                    ),
                    "mark3": st.column_config.NumberColumn(
                        "Mark 3",
                        min_value=0,
                        max_value=100,
                        step=0.5,
                    ),
                }
            )

            # Calculate Result button
            if st.button("Calculate Result"):
                result_df, tie_groups = calculate_total_and_rank(edited_df, point_values)

                # Display tie information
                display_tie_information(tie_groups)

                st.write("Results calculated! Updated data:")
                st.dataframe(result_df)

                # Display points summary
                st.subheader("Points Summary")
                points_info = f"""
                Current Points Distribution:
                - Rank 1: {point_values['first']} points
                - Rank 2: {point_values['second']} points
                - Rank 3: {point_values['third']} points
                - Other ranks: 0 points
                """
                st.info(points_info)

            # Download Result button
            if st.button("Download Result"):
                if not event_name:
                    st.error("Please enter an event name before downloading")
                    return

                # Calculate results again to ensure we have the latest data
                result_df, tie_groups = calculate_total_and_rank(edited_df, point_values)

                # If master file is uploaded, merge the data
                if master_file is not None:
                    master_df = pd.read_excel(master_file)
                    st.write("Preview of Master Data:")
                    st.dataframe(master_df.head())

                    final_df = merge_with_master_data(result_df, master_df)
                    if final_df is None:
                        return
                else:
                    final_df = result_df

                # Create Excel file in memory
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False)

                # Generate download link
                output.seek(0)
                st.download_button(
                    label="Click to Download",
                    data=output,
                    file_name=f"{event_name}_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.error("Please make sure your Excel files have the correct format")


if __name__ == "__main__":
    main()
