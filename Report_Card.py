import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import os

# Function to load data from the Excel file
def load_data(file_path):
    try:
        # Load the data into a pandas DataFrame
        df = pd.read_excel(file_path)

        # Check if required columns exist
        required_columns = {'id', 'Name', 'Gender', 'Age', 'Section', 'Science', 'English', 'History', 'Maths'}
        if not required_columns.issubset(df.columns):
            raise ValueError("The Excel file must contain the columns: id, Name, Gender, Age, Section, Science, English, History, Maths")

        # Check for missing or invalid data
        if df.isnull().any().any():
            raise ValueError("The Excel file contains missing data. Please clean the data and try again.")

        return df
    except Exception as e:
        print(f"Error loading file: {e}")
        return None

# Function to generate a report card for a student
def generate_report_card(student_data, output_directory):
    try:
        # Extract student details
        student_id = student_data['id']
        name = student_data['Name']
        gender = student_data['Gender']
        age = student_data['Age']
        section = student_data['Section']

        # Calculate scores
        subject_scores = {
            'Science': student_data['Science'],
            'English': student_data['English'],
            'History': student_data['History'],
            'Maths': student_data['Maths'],
        }
        total_score = sum(subject_scores.values())
        average_score = total_score / len(subject_scores)

        # Generate PDF
        filename = f"{output_directory}/report_card_{student_id}.pdf"
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter

        # Add student details
        c.setFont("Helvetica-Bold", 16)
        c.setFillColor(colors.darkblue)
        c.drawString(100, height - 80, f"Report Card - {name} (ID: {student_id})")
        c.setFont("Helvetica", 12)
        c.setFillColor(colors.black)
        c.drawString(100, height - 120, f"Gender: {gender}")
        c.drawString(100, height - 140, f"Age: {age}")
        c.drawString(100, height - 160, f"Section: {section}")

        # Add total and average scores
        c.drawString(100, height - 200, f"Total Score: {total_score}")
        c.drawString(100, height - 220, f"Average Score: {average_score:.2f}")

        # Add a table for subject-wise scores
        data = [['Subject', 'Score']] + [[subject, score] for subject, score in subject_scores.items()]
        table = Table(data, colWidths=[200, 100])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))

        # Position the table and draw it
        table.wrapOn(c, width, height)
        table.drawOn(c, 100, height - 320)

        # Save the PDF
        c.save()
        print(f"Report card for {name} saved at {filename}")
    except Exception as e:
        print(f"Error generating report card for ID {student_data['id']}: {e}")

# Main function to generate all report cards
def main():
    # File path to the Excel file
    file_path = '/content/student_scores.xlsx'  # Update the path here
    output_directory = 'report_cards'
    os.makedirs(output_directory, exist_ok=True)

    # Load the student data
    df = load_data(file_path)
    if df is None:
        return

    # Process each student's data
    for _, student_data in df.iterrows():
        generate_report_card(student_data, output_directory)

# Run the script
if __name__ == "__main__":
    main()
