import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO


def load_and_process_data(uploaded_file):
    """Load and process the Excel file."""
    df = pd.read_excel(uploaded_file)
    return df


def get_top_performers(df, category, n=5):
    """Get top n performers for a given category."""
    if category == 'Student':
        return df.groupby(['ID No', 'Student Name', 'Church', 'Section', 'Region'])['Points'].sum() \
            .reset_index().sort_values('Points', ascending=False).head(n)

    elif category == 'Church':
        return df.groupby('Church')['Points'].sum() \
            .reset_index().sort_values('Points', ascending=False).head(n)

    elif category == 'Section':
        return df.groupby('Section')['Points'].sum() \
            .reset_index().sort_values('Points', ascending=False).head(n)

    elif category == 'Region':
        return df.groupby('Region')['Points'].sum() \
            .reset_index().sort_values('Points', ascending=False).head(n)


def create_visualization(data, category):
    """Create bar chart visualization for top performers."""
    if category == 'Student':
        fig = px.bar(data, x='Student Name', y='Points',
                     title=f'Top 5 {category}s by Points',
                     text='Points',
                     hover_data=['Church', 'Section', 'Region'])
    else:
        fig = px.bar(data, x=category, y='Points',
                     title=f'Top 5 {category}s by Points',
                     text='Points')

    fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    return fig


def export_results_to_excel(top_performers, category, summary_stats):
    """Export results to Excel file."""
    with BytesIO() as buffer:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write top performers
            top_performers.to_excel(writer, sheet_name='Top Performers', index=False)

            # Write summary statistics
            pd.DataFrame([summary_stats]).to_excel(writer, sheet_name='Summary Statistics', index=False)

        return buffer.getvalue()


# Main Streamlit app
st.title('Competition Analysis Dashboard')

# Title input field
competition_title = st.text_input('Enter Competition Title', 'Competition Results')

# File upload
uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Load and process data
    df = load_and_process_data(uploaded_file)

    # Category selection (moved below file upload)
    category = st.selectbox('Select Category',
                            ['Student', 'Church', 'Section', 'Region'])

    # Get top performers for selected category
    top_performers = get_top_performers(df, category)

    # Display results
    st.subheader(f'Top 5 {category}s')
    st.dataframe(top_performers)

    # Create and display visualization
    fig = create_visualization(top_performers, category)
    st.plotly_chart(fig)

    # Additional statistics
    st.subheader('Summary Statistics')
    summary_stats = {}
    if category == 'Student':
        total_students = df['ID No'].nunique()
        total_churches = df['Church'].nunique()
        st.write(f'Total Students: {total_students}')
        st.write(f'Total Churches: {total_churches}')
        summary_stats = {
            'Total Students': total_students,
            'Total Churches': total_churches
        }
    elif category == 'Church':
        total_churches = df['Church'].nunique()
        avg_points = df.groupby('Church')['Points'].mean().mean()
        st.write(f'Total Churches: {total_churches}')
        st.write(f'Average Points per Church: {avg_points:.2f}')
        summary_stats = {
            'Total Churches': total_churches,
            'Average Points per Church': round(avg_points, 2)
        }
    elif category == 'Section':
        total_sections = df['Section'].nunique()
        avg_points = df.groupby('Section')['Points'].mean().mean()
        st.write(f'Total Sections: {total_sections}')
        st.write(f'Average Points per Section: {avg_points:.2f}')
        summary_stats = {
            'Total Sections': total_sections,
            'Average Points per Section': round(avg_points, 2)
        }
    else:
        total_regions = df['Region'].nunique()
        avg_points = df.groupby('Region')['Points'].mean().mean()
        st.write(f'Total Regions: {total_regions}')
        st.write(f'Average Points per Region: {avg_points:.2f}')
        summary_stats = {
            'Total Regions': total_regions,
            'Average Points per Region': round(avg_points, 2)
        }

    # Download button
    excel_data = export_results_to_excel(top_performers, category, summary_stats)
    st.download_button(
        label="Download Results",
        data=excel_data,
        file_name=f"{competition_title.replace(' ', '_')}_{category}_results.xlsx",
        mime="application/vnd.ms-excel"
    )
