import os
import win32com.client

# Define a global list to store headings with enumeration
global_headings = []

def extract_headings_with_enumeration(doc_path):
    # Access the global variable
    global global_headings

    # Check if the file exists
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"File not found at: {doc_path}")

    # Start the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # Keep Word hidden

    # Open the document
    docs = word.Documents.Open(doc_path)

    # Extract headings and enumeration
    for paragraph in docs.Paragraphs:
        if paragraph.Range.Style.NameLocal.startswith("Heading"):
            # Extract text and numbering
            text = paragraph.Range.Text.strip()
            numbering = paragraph.Range.ListFormat.ListString  # Extract enumeration
            # Save the result as a combined string to the global list
            global_headings.append(f"{numbering} {text}")

    # Ensure Word and the document are closed properly
    docs.Close(False)
    word.Quit()


# Replace with the path to your Word document
doc_path = r"C:\Users\NITRO\PycharmProjects\test\Title.docx"

# Call the function to extract headings
extract_headings_with_enumeration(doc_path)

# Print the global list
print("Global Headings List:")
for heading in global_headings:
    print(heading)
