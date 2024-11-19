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
    is_multiple_choice = False
    is_true_false = False

    title = os.path.splitext(os.path.basename(file_path))[0]

    for para in document.paragraphs:
        text = para.text.strip()

        if "Quiz Questions" in text:
            reading_questions = True
            continue

        if reading_questions:
            match = re.match(r"^(\d+)\.\s+(.*)", text)
            if match:
                question_number = int(match.group(1))
                if 1 <= question_number <= 20:
                    if current_question:
                        questions.append((current_question, answers, correct_answer, remarks, question_count, is_true_false, is_multiple_choice))
                    current_question = match.group(2)
                    answers = []
                    correct_answer = ""
                    remarks = ""
                    question_count = question_number
                    is_true_false = current_question.startswith("True or False:")
                    is_multiple_choice = current_question.startswith("Select all that apply:")
            elif text.startswith("Answer:") or text.startswith("Correct Answer:"):
                if is_multiple_choice:
                    # Split multiple correct answers into a list of letters
                    correct_answer = "|".join([f"answer{['A', 'B', 'C', 'D', 'E'].index(ans.strip()[0]) + 1}" for ans in text.split(":")[1].split(",")])
                else:
                    correct_answer = text.split(":")[1].strip()[0]  # Single answer (A, B, etc.)
                if "Rationale:" in text:
                    # Handle combined Correct Answer and Rationale in one line
                    parts = re.split(r"Rationale:", text, maxsplit=1)
                    if len(parts) > 1:
                        remarks = parts[1].strip()  # Extract the rationale                        
                    else:
                        remarks = ""               

            elif "Rationale:" in text:
                # Handle standalone Rationale
                remarks = text.split(":", 1)[1].strip()
                
            elif text.startswith("•") or re.match(r"^[A-E]\)", text):
                # Check if multiple options exist in the same line
                options = re.split(r"(?<!^)\s+(?=[A-E]\))", text)
                if len(options) > 1:  # Multiple options in the same line
                    for option in options:
                        match = re.match(r"^[A-E]\)", option)
                        if match:
                            answers.append(option[3:].strip())  # Skip 'A)', 'B)', etc.
                else:
                    # Single option, handle normally
                    if text.startswith("•"):
                        answers.append(text[2:].strip())  # Bullet point
                    elif re.match(r"^[A-E]\)", text):
                        answers.append(text[3:].strip())  # Lettered option

            # elif text.startswith("•") or (len(text) > 1 and text[0] in ['A', 'B', 'C', 'D', 'E'] and text[1] == ')'):
            #     answers.append(text[2:].strip())

        elif text:
            paragraphs.append(text)

    if current_question:
        questions.append((current_question, answers, correct_answer, remarks, question_count, is_true_false, is_multiple_choice))

    description = "<br>".join(paragraphs)
    return title, description, questions

def create_csv_files(title, paragraphs, questions):
    quiz_uuid = uuid.uuid4().hex[:8]
    questions_uuid = uuid.uuid4().hex[:8]

    generated_folder = os.path.join(os.getcwd(), "generated_files")
    os.makedirs(generated_folder, exist_ok=True)

    specific_folder = os.path.join(generated_folder, title)
    os.makedirs(specific_folder, exist_ok=True)

    metadata_filename = os.path.join(specific_folder, f"quiz_{quiz_uuid}_metadata.csv")
    df_meta = pd.DataFrame({
        "Title": [title],
        "Subtitle": [f"{title} Quiz"],
        "Description": [paragraphs]
    })
    df_meta.to_csv(metadata_filename, index=False)

    question_data = []
    for question, answers, correct_answer, remarks, order, is_true_false, is_multiple_choice in questions:
        if is_true_false:
            answers = ["TRUE", "FALSE"]
            correct_answer_mapped = "answer1" if correct_answer.upper() == "TRUE" else "answer2"
        elif is_multiple_choice:
            correct_answer_mapped = correct_answer  # Already formatted as "answer1|answer3|answer4"
        else:
            correct_answer_mapped = f"answer{['A', 'B', 'C', 'D', 'E'].index(correct_answer) + 1}" if correct_answer else ""

        answer_dict = {
            "question": question,
            "question_type": "multiple_choice" if is_multiple_choice else "true_false" if is_true_false else "single_choice",
            "answer1": answers[0] if len(answers) > 0 else "",
            "answer2": answers[1] if len(answers) > 1 else "",
            "answer3": answers[2] if len(answers) > 2 else "",
            "answer4": answers[3] if len(answers) > 3 else "",
            "answer5": answers[4] if len(answers) > 4 else "",
            "correct_answer": correct_answer_mapped,
            "score": 1,
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

root = Tk()
root.title("Word to CSV Converter")
root.geometry("400x200")

label_instructions = Label(root, text="Select Word document(s) to generate CSV files:")
label_instructions.pack(pady=20)

button_select = Button(root, text="Select Files", command=open_file_dialog)
button_select.pack(pady=10)

label_status = Label(root, text="")
label_status.pack(pady=10)

root.mainloop()
