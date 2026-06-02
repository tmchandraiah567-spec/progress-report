# Student Result Management System

A Flask-based student result management application for recording exam scores, calculating grades, and reviewing student performance.

## Key Features

- Student submission form with dynamic subject input and marks entry
- USN validation and duplicate entry prevention
- Faculty login area with secure dashboard access
- Filter results by semester, exam, branch, and year
- Top N student selection and subject-wise topper ranking
- Edit, delete, and clear student records
- Results stored in `students.xlsx` automatically using `pandas` and `openpyxl`

## Requirements

- Python
- Flask
- pandas
- openpyxl

Install required packages with:

```bash
pip install -r requirements.txt
```

## Project Structure

- `app.py` - main Flask application
- `requirements.txt` - Python dependencies
- `templates/` - HTML templates for home, student, faculty, login, edit, and thank-you pages
- `static/style.css` - application styling
- `students.xlsx` - Excel file used to store student records
- `ngrok/` - contains `ngrok.exe` for tunneling the local app if needed

## Usage

1. Run the application:

```bash
python app.py
```

2. Open a browser and navigate to:

```text
http://127.0.0.1:5000/
```

3. Use the student form to submit exam details and marks.
4. Access the faculty dashboard:

```text
http://127.0.0.1:5000/faculty_login
```

5. Login credentials:

- Username: `admin`
- Password: `admin123`

## Notes

- `students.xlsx` is created automatically if it does not exist.
- Each subject entry and mark is stored as JSON inside the Excel file.
- The app supports multiple semesters, branches, and exam sessions.

## ngrok (Optional)

If you want to expose the local server externally, use the `ngrok.exe` in the `ngrok/` folder:

```bash
cd ngrok
.\ngrok.exe http 5000
```
