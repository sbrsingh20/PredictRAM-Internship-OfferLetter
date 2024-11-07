import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from datetime import date

# Load candidates
candidates_file = 'candidates.xlsx'  # Update this path as necessary
candidates_df = pd.read_excel(candidates_file)

# Check for leading/trailing spaces in column names
candidates_df.columns = candidates_df.columns.str.strip()

# Function to generate letter of recommendation
def generate_lor(name, start_date, end_date):
    pdf_file = f"LOR_{name}.pdf"
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter
    
    # Add border
    c.setStrokeColor(colors.navy)
    c.setLineWidth(2)
    c.rect(0.5 * inch, 0.5 * inch, width - inch, height - inch)

    # Add logo image (Replace 'image.png' with the actual path to your logo)
    logo_path = 'image.png'  # Path to your logo
    c.drawImage(logo_path, (width - 200) / 2, height - 140, 200, 100)  # Center the logo

    # Set today's date at the top
    today_date = date.today().strftime('%d-%m-%Y')
    c.setFont("Helvetica", 12)
    c.drawString(0.75 * inch, height - 170, "Issue Date: " + today_date)

    # Center and style "To Whom It May Concern"
    c.setFont("Helvetica-Bold", 16)
    c.setFillColor(colors.blue)
    c.drawCentredString(width / 2, height - 200, "To Whom It May Concern")
    
    # Reset font color to black for the rest of the text
    c.setFillColor(colors.black)

    # Add letter content
    c.setFont("Helvetica", 12)
    content_y_start = height - 240  # Start drawing text below the salutation
    c.drawString(0.75 * inch, content_y_start, f"It gives me great pleasure to write this letter of recommendation for {name},")
    c.drawString(0.75 * inch, content_y_start - 20, f"I have worked side-by-side with {name} from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')} in the")
    c.drawString(0.75 * inch, content_y_start - 40, "financial analysis & equity research department at PredictRAM. During tenure,")
    c.drawString(0.75 * inch, content_y_start - 60, f"{name} served as an Equity Research (intern) and was a direct report to me.")
    c.drawString(0.75 * inch, content_y_start - 80, f"As a direct report, {name} was a successful, easy-to-manage associate, and always gave")
    c.drawString(0.75 * inch, content_y_start - 100, "that extra effort to meet deadlines. Demonstrated superior analytical capabilities")
    c.drawString(0.75 * inch, content_y_start - 120, "and soon became an expert in financial analysis & equity research.")
    c.drawString(0.75 * inch, content_y_start - 160, f"During the internship with our organization, I have experienced an individual who shows up")
    c.drawString(0.75 * inch, content_y_start - 180, "earlier than asked, works hard, and carries themselves in a polite, respectable manner.")
    c.drawString(0.75 * inch, content_y_start - 200, f"Intern is exceptionally brilliant, has consistently shown strong leadership qualities.")
    c.drawString(0.75 * inch, content_y_start - 220, f"Intern is indeed a talented and dedicated student.")
    c.drawString(0.75 * inch, content_y_start - 240, f"I would recommend {name} for the Financial Analyst role.")
    c.drawString(0.75 * inch, content_y_start - 260, "I wish all the luck for academic success.")

    c.drawString(0.75 * inch, content_y_start - 280, f"Duration of Internship from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')} .")    
    c.drawString(0.75 * inch, content_y_start - 300, "Sincerely,")
    
    # Footer with signature and address
    y = content_y_start - 400
    c.drawImage('SignWithLogo.png', 0.75 * inch, y, 190, 80)  # Replace with actual path
    c.drawString(0.75 * inch, y - 50, "Subir Singh")
    c.drawString(0.75 * inch, y - 65, "Director- PredictRAM(Params Data provider Pvt Ltd")
    c.drawString(0.75 * inch, y - 80, "Office: B1/638 A, 2nd & 3rd Floor, Janakpuri New Delhi 110058")
    c.drawString(0.75 * inch, y - 95, "")
    
    c.showPage()
    c.save()
    return pdf_file

# Streamlit App

 # Add a logo to the top header
st.image("png_2.3-removebg-preview.png", width=400)  # Replace "your_logo.png" with the path to your logo
st.title("Letter of Recommendation Generator")

# Input Form
st.header("Enter LOR Details")

with st.form("lor_form"):
    name = st.text_input("Candidate Name")
    start_date = st.date_input("Start Date")
    end_date = st.date_input("End Date")
    submit_button = st.form_submit_button(label="Generate LOR")

# Check if the candidate is in the list
if submit_button:
    if 'Candidate Name' in candidates_df.columns:
        if name in candidates_df['Candidate Name'].values:
            # Generate LOR PDF
            pdf_file = generate_lor(name, start_date, end_date)
            
            with open(pdf_file, "rb") as file:
                st.download_button(
                    label="Download LOR",
                    data=file,
                    file_name=pdf_file,
                    mime="application/pdf"
                )
        else:
            st.error("Candidate not found in the list.")
    else:
        st.error("'Candidate Name' column not found in the candidates file.")
