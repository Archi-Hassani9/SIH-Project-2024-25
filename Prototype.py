import streamlit as st
import pandas as pd
from scholarly import scholarly, MaxTriesExceededException
from docx import Document
import bibtexparser
from io import BytesIO
import time

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

@st.cache_data
def readdf(file):
    fn = file.name
    if fn.endswith('.bib'):
        content = file.read().decode('utf-8')
        bib_db = bibtexparser.load(io.StringIO(content))
        return pd.DataFrame(bib_db.entries)
    elif fn.endswith('.xlsx'):
        return pd.read_excel(file)
    else:
        st.error("Unsupported file format. Please upload a .bib or .xlsx file.")
        return None

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

@st.cache_data
def getPub(authorName):
    query = scholarly.search_author(authorName)
    for _ in range(3):
        try:
            author = scholarly.fill(next(query))
            break
        except (StopIteration, MaxTriesExceededException):
            time.sleep(5)
    else:
        st.error("Failed to fetch data from Google Scholar. Try again later.")
        return pd.DataFrame()

    pubs = [
        {
            'title': pub.get('bib', {}).get('title', ''),
            'year': pub.get('bib', {}).get('pub_year', ''),
            'journal': pub.get('bib', {}).get('venue', ''),
            'authors': pub.get('bib', {}).get('author', '')
        }
        for pub in author['publications']
    ]
    return pd.DataFrame(pubs)

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def filByYear(data, start, end):
    data['year'] = pd.to_numeric(data['year'], errors='coerce')
    return data[(data['year'] >= start) & (data['year'] <= end)]

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def saveExcel(data, filename):
    excel_stream = BytesIO()

    with pd.ExcelWriter(excel_stream, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name='Sheet1')

    excel_stream.seek(0)

    st.download_button(
        label="Download Excel File",
        data=excel_stream,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_button_excel_unique_key"
    )
    st.success(f"Data saved to {filename}. You can download it using the link above.")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def saveWord(data, filename):
    doc = Document()
    doc.add_heading('Publication Summary', 0)

    for _, row in data.iterrows():
        doc.add_paragraph(f"{row['year']}: {row['title']} ({row['journal']})")

    doc_stream = BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)

    st.download_button(
        label="Download Publication Summary",
        data=doc_stream,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="download_button_unique_key"
    )

    st.success("You can download the file using the link above.")

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

def main():
    st.title("Faculty Publication Summary Tool")
    df = st.sidebar.file_uploader("Upload a BibTeX or Excel file", type=["bib", "xlsx"])

    if df:
        data = readdf(df)
        if data is not None:
            st.write("### Uploaded Data", data.head())
            if 'authorName' not in st.session_state:
                st.session_state.authorName = "Jane Doe"
            authorName = st.text_input("Enter Author Name", st.session_state.authorName)

            searchOpt = st.radio(
                "Choose search option:",
                ('Search within the uploaded dataset', 'Search universally using Google Scholar')
            )

            if st.button("Get Publications"):
                st.session_state.authorName = authorName
                if searchOpt == 'Search within the uploaded dataset':
                    publications = data[data['author'].str.contains(authorName, case=False, na=False)]
                    if not publications.empty:
                        st.session_state.publications = publications
                        st.write(f"### Publications for {authorName} (from uploaded dataset)", publications)
                    else:
                        st.warning(f"No publications found for {authorName} in the uploaded dataset")
                else:
                    with st.spinner('Fetching publications...'):
                        publications = getPub(authorName)
                    if not publications.empty:
                        st.session_state.publications = publications
                        st.write(f"### Publications for {authorName} (from Google Scholar)", publications)
                    else:
                        st.warning(f"No publications found for {authorName}")

            if 'publications' in st.session_state:
                year1 = st.slider("Start Year", 1900, 2024, 2015)
                year2 = st.slider("End Year", 1900, 2024, 2020)
                filtered_data = filByYear(st.session_state.publications, year1, year2)
                st.write(f"### Filtered Publications ({year1}-{year2})", filtered_data)

                if st.button("Save to Excel"):
                    saveExcel(filtered_data, 'publication_summary.xlsx')
                if st.button("Save to Word"):
                    saveWord(filtered_data, 'publication_summary.docx')

if __name__ == "__main__":
    main()
