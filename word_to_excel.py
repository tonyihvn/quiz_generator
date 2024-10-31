import os
import uuid
import pandas as pd
from tkinter import Tk, Button, Label, filedialog
from docx import Document
import re

def extract_content_from_word(file_path):
    document = Document(file_path)
    paragraphs = []
    questions = []
    current_question = ""
    answers = []
    correct_answer = ""
    remarks = ""
    reading_questions = False
    question_count = 0

    # Extract file name for title
    title = os.path.splitext(os.path.basename(file_path))[0]

    for para in document.paragraphs:
        text = para.text.strip()
        
        # Start reading after "Quiz Questions"
        if "Quiz Questions" in text:
            reading_questions = True
            continue  # Skip this line

        if reading_questions:
            # Check for question numbers (supporting 1. to 20. using regular expression)
            match = re.match(r"^(\d+)\.\s+(.*)", text)
            if match:
                question_number = int(match.group(1))
                if 1 <= question_number <= 20:  # Ensure question number is within range
                    if current_question:  # Add the previous question if it's there
                        questions.append((current_question, answers, correct_answer, remarks, question_count))
                    current_question = match.group(2)  # Get the question text after the number and dot
                    answers = []  # Reset answers for new question
                    correct_answer = ""  # Reset correct answer flag
                    remarks = ""  # Reset remarks flag
                    question_count = question_number  # Set the question counter based on actual question number
            elif text.startswith("Correct Answer:"):
                correct_answer = text.split(":")[-1].strip()
            elif text.startswith("Rationale:"):
                remarks = text.split(":")[-1].strip()  # Capture remarks from Rationale
            elif text.startswith("•"):  # Detect bullet points for answers
                answers.append(text[2:].strip())  # Skip the bullet character and space
            elif text and len(text) > 1 and text[0] in ['A', 'B', 'C', 'D', 'E'] and text[1] == ')':
                answers.append(text[2:].strip())  # Capture the answer text for A), B), C), D)

        elif text:  # Collect other text as paragraphs
            paragraphs.append(text)

    # Add the last question if any
    if current_question:
        questions.append((current_question, answers, correct_answer, remarks, question_count))

    # Convert paragraphs to HTML format
    description = "<br>".join(paragraphs)  # Simple line break for HTML

    return title, description, questions

def create_csv_files(title, paragraphs, questions):
    # Generate unique identifiers
    quiz_uuid = uuid.uuid4().hex[:8]
    questions_uuid = uuid.uuid4().hex[:8]

    # Create a directory for generated files if it doesn't exist
    generated_folder = os.path.join(os.getcwd(), "generated_files")
    os.makedirs(generated_folder, exist_ok=True)

    # Create a subdirectory for this run using the quiz title
    specific_folder = os.path.join(generated_folder, title)
    os.makedirs(specific_folder, exist_ok=True)

    # First CSV file - Metadata
    metadata_filename = os.path.join(specific_folder, f"quiz_{quiz_uuid}_metadata.csv")
    df_meta = pd.DataFrame({
        "Title": [title],
        "Subtitle": [f"{title} Quiz"],
        "Description": [paragraphs]
    })
    df_meta.to_csv(metadata_filename, index=False)

    # Second CSV file - Questions
    question_data = []
    for question, answers, correct_answer, remarks, order in questions:
        # Map correct_answer to the correct answer field
        correct_answer_mapped = ""
        if correct_answer == "A":
            correct_answer_mapped = "answers1"
        elif correct_answer == "B":
            correct_answer_mapped = "answers2"
        elif correct_answer == "C":
            correct_answer_mapped = "answers3"
        elif correct_answer == "D":
            correct_answer_mapped = "answers4"
        elif correct_answer == "E":
            correct_answer_mapped = "answers5"

        # Ensure all required columns are included, even if empty
        answer_dict = {
            "question": question,
            "question_type": "single_choice",  # Assuming question type is multiple-choice
            "answer1": answers[0] if len(answers) > 0 else "",
            "answer2": answers[1] if len(answers) > 1 else "",
            "answer3": answers[2] if len(answers) > 2 else "",
            "answer4": answers[3] if len(answers) > 3 else "",
            "answer5": answers[4] if len(answers) > 4 else "",
            "correct_answer": correct_answer_mapped,  # Mapped correct answer
            "score": "",  # Add an empty score field if unspecified
            "remarks": remarks,
            "ordering": order
        }
        question_data.append(answer_dict)

    questions_filename = os.path.join(specific_folder, f"questions_{questions_uuid}.csv")
    df_questions = pd.DataFrame(question_data)
    df_questions.to_csv(questions_filename, index=False)

    return metadata_filename, questions_filename

def process_files(file_paths):
    for file_path in file_paths:
        title, paragraphs, questions = extract_content_from_word(file_path)
        metadata_filename, questions_filename = create_csv_files(title, paragraphs, questions)
        print(f"Generated files: {metadata_filename}, {questions_filename}")

def open_file_dialog():
    file_paths = filedialog.askopenfilenames(title="Select Word Files", filetypes=[("Word files", "*.docx")])
    if file_paths:
        process_files(file_paths)
        label_status.config(text="CSV files generated successfully!")
    else:
        label_status.config(text="No files selected.")

# Set up the GUI
root = Tk()
root.title("Word to CSV Converter")
root.geometry("400x200")

# Label and Button
label_instructions = Label(root, text="Select Word document(s) to generate CSV files:")
label_instructions.pack(pady=20)

button_select = Button(root, text="Select Files", command=open_file_dialog)
button_select.pack(pady=10)

label_status = Label(root, text="")
label_status.pack(pady=10)

root.mainloop()
