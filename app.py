import streamlit as st
import openpyxl
import io

# Streamlit app title
st.title("Excel Row keep and Delete App")

# Upload workbook1 (the main workbook)
workbook1_file = st.file_uploader("Upload the main workbook (keyword test)", type=['xlsx'])

# Upload workbook2 (the workbook with words to delete and keep)
workbook2_file = st.file_uploader("Upload the workbook with words to delete and keep (Keep and delete)", type=['xlsx'])

if workbook1_file and workbook2_file:
    # Load the workbooks
    workbook1 = openpyxl.load_workbook(workbook1_file)
    workbook2 = openpyxl.load_workbook(workbook2_file)

    # Get the sheets
    sheet1 = workbook1["Sheet1"]
    sheet2_words = workbook2["Delete"]
    sheet2_exceptions = workbook2["Keep"]

    # Extract words to delete and words to keep
    words_to_delete = set([cell.value.lower() for row in sheet2_words.iter_rows() for cell in row if cell.value is not None])
    words_to_keep = set([cell.value.lower() for row in sheet2_exceptions.iter_rows() for cell in row if cell.value is not None])

    # Iterate through rows in reverse order to safely delete
    for row_index in range(sheet1.max_row, 0, -1):
        row = sheet1[row_index]
        row_text = " ".join([str(cell.value).lower() if cell.value is not None else "" for cell in row])

        # Check if any word to delete is present and no word to keep is present
        if any(word in row_text for word in words_to_delete) and not any(word in row_text for word in words_to_keep):
            sheet1.delete_rows(row_index)

    # Save the modified workbook to an in-memory file
    output = io.BytesIO()
    workbook1.save(output)
    output.seek(0)

    # Provide the download link for the updated file
    st.success("Rows deleted successfully!")
    st.download_button(
        label="Download the updated workbook",
        data=output,
        file_name="keyword_test_updated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload both Excel files to proceed.")
