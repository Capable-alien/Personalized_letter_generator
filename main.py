import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
import io

# Streamlit App
st.title("Personalized Letter Generator")

# Upload CSV or Excel file
uploaded_file = st.file_uploader("Choose a file", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # Read the file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.write("Data Preview:")
        st.dataframe(df.head())

        # User inputs
        from_name = st.text_input("From (Your Name)")
        area_street = st.text_area("From (Area/Street)")
        city_town = st.text_input("From (City/Town)")
        state_country = st.text_input("From (State/Country)")
        pincode = st.text_input("From (Pincode)")
        date_input = st.date_input("Date", datetime.today())
        letter_body = st.text_area("Body of the letter")

        # Check if all inputs are provided
        if st.button("Generate Letters"):
            if not from_name or not area_street or not city_town or not state_country or not pincode or not letter_body:
                st.warning("Please fill in all the fields: From Name, From Address (Area/Street, City/Town, State/Country, Pincode), and Body of the Letter.")
            else:
                # Function to generate letter for a student
                def generate_letter(student):
                    try:
                        # Define letter template
                        letter_template = """
                        From: {from_name}
                        {area_street},
                        {city_town}, {state_country}, {pincode}.

                        Date: {date}

                        To:
                            {name}
                            {city}
                            {email}
                            {phone}

                            
                        Dear {name},

                        {body}

                        
                        
                        Sincerely,
                        {from_name}
                        """

                        return letter_template.format(
                            from_name=from_name,
                            area_street=area_street,
                            city_town=city_town,
                            state_country=state_country,
                            pincode=pincode,
                            date=date_input.strftime('%B %d, %Y'),
                            name=student['name'],
                            city=student['city'],
                            email=student['email'],
                            phone=student['phone'],
                            body=letter_body
                        )
                    except Exception as e:
                        st.error(f"An error occurred while generating the letter: {e}")

                # Generate letters for each student
                for _, student in df.iterrows():
                    letter_content = generate_letter(student)

                    # Generate PDF
                    if letter_content:
                        try:
                            pdf_buffer = io.BytesIO()
                            c = canvas.Canvas(pdf_buffer, pagesize=letter)
                            width, height = letter
                            for i, line in enumerate(letter_content.split('\n')):
                                c.drawString(72, height - 72 - 14 * i, line)
                            c.save()

                            # Save PDF
                            pdf_buffer.seek(0)
                            st.download_button(
                                label=f"Download Letter for {student['name']}",
                                data=pdf_buffer,
                                file_name=f"letter_{student['name']}.pdf",  # Use name instead of rollno
                                mime='application/pdf'
                            )
                        except Exception as e:
                            st.error(f"An error occurred while saving PDF for {student['name']}: {e}")

    except Exception as e:
        st.error(f"An error occurred while reading the file: {e}")

else:
    st.info("Please upload a CSV or Excel file containing student details.")
