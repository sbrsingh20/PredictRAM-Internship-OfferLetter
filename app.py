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
    c.setLineWidth(1)
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
    c.drawCentredString(width / 2, height - 200, "Internship Offer Letter")
    
    # Reset font color to black for the rest of the text
    c.setFillColor(colors.black)

    # Add letter content
    c.setFont("Helvetica", 12)
    content_y_start = height - 240  # Start drawing text below the salutation
    c.drawString(0.75 * inch, content_y_start, f"Dear {name},")
    c.drawString(0.75 * inch, content_y_start - 40, f"I am delighted & excited to welcome you to PredictRAM for Financial Analyst Internship")
    c.drawString(0.75 * inch, content_y_start - 60, "At PredictRAM, we believe that our team is our biggest strength, and we take pride in hiring only")
    c.drawString(0.75 * inch, content_y_start - 80, f"the best and the brightest. We are confident that you would play a significant role in the overall ")
    c.drawString(0.75 * inch, content_y_start - 100, f"success of the venture and wish you the most enjoyable, learning-packed, and truly meaningful ")
    c.drawString(0.75 * inch, content_y_start - 120, "internship experience with PredictRAM.")
    c.drawString(0.75 * inch, content_y_start - 160, f"Your appointment will be governed by the terms and conditions presented in")
    c.drawString(0.75 * inch, content_y_start - 180, "https://predictram.gitbook.io/docs/internship-terms")
    c.drawString(0.75 * inch, content_y_start - 220, f"We look forward to you joining us. Please do not hesitate to call us for any information. ")
    c.drawString(0.75 * inch, content_y_start - 240, f"Also, please sign the duplicate of this offer as your acceptance and forward the same to us.")
    c.drawString(0.75 * inch, content_y_start - 260, f"Congratulations! {name}")
    c.drawString(0.75 * inch, content_y_start - 280, "")

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
