from flask import Flask, render_template, request, redirect, session, jsonify
import pandas as pd
import os
import json

app = Flask(__name__)
app.secret_key = "secret123"

# Add custom Jinja2 filter to parse JSON in templates
@app.template_filter('from_json')
def from_json_filter(json_str):
    try:
        return json.loads(json_str)
    except:
        return []

# Add enumerate to Jinja2 globals
app.jinja_env.globals.update(enumerate=enumerate)

FILE = "students.xlsx"
# Dynamic subjects - now input by users instead of predefined
SEMESTERS = [str(i) for i in range(1, 9)]
YEARS = ['2025', '2026', '2027']
BRANCHES = ['CSE', 'ISE', 'EEE', 'ECE', 'CSE-CC', 'CSE-BC BS', 'CSE- CY', 'CSE- IY']
EXAMS = ['IA-1', 'IA-2', 'IA-3']
COLUMNS = [
    'exam', 'semester', 'branch', 'year', 'usn', 'name',
    'subjects', 'marks',
    'total', 'average', 'grade'
]

# Preferred subject display order (case-insensitive match)
SUBJECT_ORDER = ['maths', 'physics', 'ds', 'python', 'fcn', 'fdbms']

# Map common variants to canonical subject keys
SUBJECT_ALIASES = {
    'math': 'maths', 'mathematics': 'maths', 'maths': 'maths',
    'physics': 'physics', 'phys': 'physics',
    'dsa': 'ds', 'data structures': 'ds', 'data-structures': 'ds', 'ds': 'ds',
    'python': 'python', 'py': 'python',
    'fcn': 'fcn',
    'fdbms': 'fdbms', 'dbms': 'fdbms', 'database': 'fdbms', 'db': 'fdbms'
}


def normalize_subject(name):
    if not name:
        return ''
    s = str(name).strip().lower()
    return SUBJECT_ALIASES.get(s, s)

if not os.path.exists(FILE):
    df = pd.DataFrame(columns=COLUMNS)
    try:
        df.to_excel(FILE, index=False)
    except PermissionError as exc:
        raise RuntimeError(
            "Cannot create students.xlsx. Close the file if it is open in another program or fix permissions."
        ) from exc


def init_df():
    df = pd.read_excel(FILE, engine='openpyxl')
    for col in COLUMNS:
        if col not in df.columns:
            if col in ['exam', 'semester', 'branch', 'year', 'usn', 'name', 'grade', 'subjects', 'marks']:
                df[col] = ""
            else:
                df[col] = 0
    return df[COLUMNS]


def load_df():
    try:
        df = init_df()
    except PermissionError as exc:
        raise RuntimeError(
            "Cannot access students.xlsx. Close the file if it is open in another program or fix permissions."
        ) from exc

    if 'semester' in df.columns:
        df['semester'] = (
            df['semester']
            .fillna('')
            .astype(str)
            .str.replace(r'\.0+$', '', regex=True)
            .str.strip()
        )
    if 'branch' in df.columns:
        df['branch'] = df['branch'].fillna('').astype(str).str.strip()
    if 'year' in df.columns:
        df['year'] = (
            df['year']
            .fillna('')
            .astype(str)
            .str.replace(r'\.0+$', '', regex=True)
            .str.strip()
        )
    if 'exam' in df.columns:
        df['exam'] = df['exam'].fillna('').astype(str).str.strip()
    return df


def save_df(df):
    try:
        df.to_excel(FILE, index=False)
        # Add filters to the Excel sheet
        from openpyxl import load_workbook
        wb = load_workbook(FILE)
        ws = wb.active
        # Add AutoFilter to the header row
        if len(df) > 0:
            ws.auto_filter.ref = f"A1:{chr(64 + len(df.columns))}{len(df) + 1}"
        wb.save(FILE)
    except PermissionError as exc:
        raise RuntimeError(
            "Cannot write students.xlsx. Close the file if it is open in another program or fix permissions."
        ) from exc


def extract_usn_number(usn):
    """Extract last 3 digits of USN as integer for sorting"""
    try:
        return int(str(usn)[-3:])
    except:
        return 999

def validate_usn(usn):
    """Validate USN format - last 3 digits should be 001 to 060"""
    usn = usn.strip().upper()
    if len(usn) < 3:
        return False, "USN must be at least 3 characters long"
    
    last_three = usn[-3:]
    if not last_three.isdigit():
        return False, "Last 3 digits of USN must be numbers"
    
    num = int(last_three)
    if num < 1 or num > 60:
        return False, "Last 3 digits must be between 001 and 060"
    
    return True, "Valid"

def grade(marks):
    if any(m < 9 for m in marks):  # < 9 out of 25 is fail
        return "Fail"
    scale = 25  # Now marks are out of 25
    avg = sum(marks) / len(marks) / scale * 100
    if avg >= 85:
        return "Distinction"
    elif avg >= 70:
        return "First Class"
    elif avg >= 50:
        return "Pass"
    else:
        return "Fail"

def get_subject_wise_toppers(filtered_df, rank_style='standard'):
    """Get top 5 students per subject from filtered dataframe and compute ranks.
    rank_style: 'standard' => 1,2,2,4 (standard competition)
                'dense'    => 1,2,2,3 (dense ranking)
    Returns dict: subject -> list of dicts with keys: usn,name,mark,average,grade,color,rank
    """
    subject_toppers = {}

    if len(filtered_df) == 0:
        return subject_toppers

    # Collect all subjects from filtered data (normalize aliases)
    all_subjects = set()
    for subjects_json in filtered_df['subjects']:
        try:
            subjects = json.loads(subjects_json)
            for subj in subjects:
                canon = normalize_subject(subj)
                if canon:
                    all_subjects.add(canon)
        except:
            pass

    # For each subject, get students and compute ranks
    for subject in sorted(all_subjects):
        students = []
        for idx, row in filtered_df.iterrows():
            try:
                subjects = json.loads(row['subjects'])
                marks = json.loads(row['marks'])
                # find the index where normalized subject matches canonical subject
                subject_idx = None
                for i, sraw in enumerate(subjects):
                    if normalize_subject(sraw) == subject:
                        subject_idx = i
                        break
                if subject_idx is None:
                    continue
                mark = marks[subject_idx]
                # Determine color based on mark
                if mark >= 20:
                    color = '#4CAF50'
                elif mark >= 15:
                    color = '#FFC107'
                else:
                    color = '#FF9800'

                students.append({
                    'usn': row['usn'],
                    'name': row['name'],
                    'mark': mark,
                    'average': row['average'],
                    'grade': row['grade'],
                    'color': color
                })
            except:
                pass

        # Sort by mark descending, then by USN ascending for stable ordering of ties
        students.sort(key=lambda x: (-x['mark'], x['usn']))

        # Compute ranks according to chosen style
        ranked = []
        prev_mark = None
        prev_rank = 0
        position = 0  # 1-based position in sorted list
        dense_rank = 0
        for s in students:
            position += 1
            if prev_mark is None or s['mark'] != prev_mark:
                # New mark
                if rank_style == 'dense':
                    dense_rank += 1
                    rank = dense_rank
                else:
                    # standard competition: rank equals current position
                    rank = position
                prev_mark = s['mark']
                prev_rank = rank
            else:
                # same mark as previous
                rank = prev_rank

            s_with_rank = s.copy()
            s_with_rank['rank'] = rank
            ranked.append(s_with_rank)

        subject_toppers[subject] = ranked[:5]

    return subject_toppers

@app.route('/')
def home():
    return render_template("home.html")

@app.route('/student', methods=['GET','POST'])
def student():
    if request.method == 'POST':
        exam = request.form.get('exam', EXAMS[0])
        semester = request.form.get('semester', SEMESTERS[0])
        branch = request.form.get('branch', '').strip()
        year = request.form.get('year', '').strip()
        usn = request.form.get('usn', '').strip().upper()
        name = request.form.get('name', '')
        
        # Validate USN format
        is_valid, msg = validate_usn(usn)
        if not is_valid:
            return render_template(
                "student.html",
                exams=EXAMS,
                semesters=SEMESTERS,
                branches=BRANCHES,
                years=YEARS,
                error=msg
            )
        
        # Get dynamic subjects and marks
        subjects = []
        marks = []
        for i in range(1, 7):
            subject = request.form.get(f'sub{i}', '').strip()
            mark = int(request.form.get(f'm{i}', '0'))
            if subject:  # Only add if subject name is provided
                subjects.append(subject)
                marks.append(mark)

        df = load_df()
        # Check for duplicate USN in same semester and year
        if not df[(df['usn'] == usn) & (df['semester'] == semester) & (df['year'] == year)].empty:
            return render_template(
                "student.html",
                exams=EXAMS,
                semesters=SEMESTERS,
                branches=BRANCHES,
                years=YEARS,
                error="Entry already exists for this USN in the selected semester and year."
            )

        total = sum(marks)
        scale = 25 * len(marks)  # Each subject is out of 25
        avg = total / scale * 100
        g = grade(marks)

        # Store subjects as JSON string
        df.loc[len(df)] = [exam, semester, branch, year, usn, name, json.dumps(subjects), json.dumps(marks), total, avg, g]
        save_df(df)

        return redirect('/thankyou')

    return render_template("student.html", exams=EXAMS, semesters=SEMESTERS, branches=BRANCHES, years=YEARS)

@app.route('/faculty_login', methods=['GET','POST'])
def faculty_login():
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        if username == "admin" and password == "admin123":
            session['logged_in'] = True
            return redirect('/faculty')
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route('/thankyou')
def thankyou():
    return render_template("thankyou.html")

@app.route('/faculty', methods=['GET', 'POST'])
def faculty():
    if not session.get('logged_in'):
        return redirect('/faculty_login')

    semester_selected = request.values.get('semester', '')
    exam_selected = request.values.get('exam', '')
    branch_selected = request.values.get('branch', '')
    year_selected = request.values.get('year', '')
    top_n = request.values.get('top_n', '1')  # New parameter for top N selection
    rank_style = request.values.get('rank_style', 'standard')  # ranking style: 'standard' or 'dense'

    df = load_df().fillna("")
    filtered = pd.DataFrame(columns=df.columns)
    filters_applied = False

    overall_avg = None
    toppers = []
    subject_avg = {}
    subject_toppers = {}
    
    if semester_selected and exam_selected and branch_selected and year_selected:
        filters_applied = True
        filtered = df[(df['semester'] == semester_selected) & (df['exam'] == exam_selected) & (df['branch'] == branch_selected) & (df['year'] == year_selected)]
        
        # Sort by USN (last 3 digits) in ascending order
        filtered['usn_number'] = filtered['usn'].apply(extract_usn_number)
        filtered = filtered.sort_values(by='usn_number', ascending=True, kind='stable')
        filtered = filtered.drop('usn_number', axis=1)

        if len(filtered) > 0:
            overall_avg = float(filtered['average'].mean())
            
            # Build header names for the subject columns from the first record
            subject_column_names = []
            try:
                subject_column_names = json.loads(filtered.iloc[0]['subjects'])
            except:
                subject_column_names = []

            # Get top N students - sort by average descending
            n = int(top_n)
            top_students = filtered.sort_values(by='average', ascending=False).head(n)
            toppers = top_students.to_dict(orient='records')
            
            # Get subject-wise top 5 students (with ranks)
            subject_toppers = get_subject_wise_toppers(filtered, rank_style=rank_style)

            # Prepare chart data per subject for frontend charts
            subject_chart_list = []
            # Order subjects according to SUBJECT_ORDER first, then any remaining
            all_subjects = list(subject_toppers.keys())
            ordered = []
            lowered = {s.lower(): s for s in all_subjects}
            for so in SUBJECT_ORDER:
                if so in lowered:
                    ordered.append(lowered[so])
            # add any subjects not in SUBJECT_ORDER at the end (sorted)
            for s in sorted(all_subjects):
                if s not in ordered:
                    ordered.append(s)

            for subj in ordered:
                lst = subject_toppers.get(subj, [])
                labels = [s['name'] for s in lst]
                marks = [s['mark'] for s in lst]
                colors = [s.get('color', '#4CAF50') for s in lst]
                # For display, keep canonical subject as requested (use provided SUBJECT_ORDER names)
                display_name = subj if subj in SUBJECT_ORDER else subj.title()
                subject_chart_list.append({
                    'subject': display_name,
                    'labels': labels,
                    'marks': marks,
                    'colors': colors
                })

        # Keep original indices for /student_details/<idx>
        # Include index in the records for template use
        data = filtered.reset_index().rename(columns={'index': 'df_index'}).to_dict(orient='records')
    else:
        data = filtered.to_dict(orient='records')

    return render_template(
        "faculty.html",
        data=data,
        toppers=toppers,
        overall_avg=overall_avg,
        semesters=SEMESTERS,
        exams=EXAMS,
        branches=BRANCHES,
        years=YEARS,
        semester_selected=semester_selected,
        exam_selected=exam_selected,
        branch_selected=branch_selected,
        year_selected=year_selected,
        filters_applied=filters_applied,
        total_records=len(filtered),
        top_n=top_n,
        rank_style=rank_style,
        subject_avg=subject_avg,
        subject_toppers=subject_toppers,
        subject_chart_list=locals().get('subject_chart_list', []),
        subject_column_names=locals().get('subject_column_names', [])
    )

@app.route('/delete/<int:i>')
def delete(i):
    df = load_df()
    df = df.drop(i).reset_index(drop=True)
    save_df(df)
    return redirect('/faculty')

# 🔥 CLEAR ALL DATA
@app.route('/clear')
def clear():
    df = pd.DataFrame(columns=COLUMNS)
    save_df(df)
    return redirect('/faculty')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect('/')


@app.route('/student_details/<int:idx>')
def student_details(idx):
    df = load_df().fillna("")
    if idx in df.index:
        row = df.loc[idx]
        try:
            subjects = json.loads(row['subjects'])
            marks = json.loads(row['marks'])
        except:
            subjects = []
            marks = []
        
        student_data = {
            'exam': str(row['exam']),
            'semester': str(row['semester']),
            'branch': str(row['branch']),
            'year': str(row['year']),
            'usn': str(row['usn']),
            'name': str(row['name']),
            'subjects': subjects,
            'marks': marks,
            'total': int(row['total']),
            'average': float(row['average']),
            'grade': str(row['grade'])
        }
        return jsonify(student_data)
    return jsonify({'error': 'Student not found'}), 404

@app.route('/edit/<int:i>', methods=['GET', 'POST'])
def edit(i):
    if not session.get('logged_in'):
        return redirect('/faculty_login')
    
    df = load_df().fillna("")
    if i not in df.index:
        return redirect('/faculty')
    
    if request.method == 'POST':
        exam = request.form.get('exam', EXAMS[0])
        semester = request.form.get('semester', SEMESTERS[0])
        branch = request.form.get('branch', '').strip()
        year = request.form.get('year', '').strip()
        usn = request.form.get('usn', '').strip().upper()
        name = request.form.get('name', '')
        
        # Get dynamic subjects and marks
        subjects = []
        marks = []
        for j in range(1, 7):
            subject = request.form.get(f'sub{j}', '').strip()
            mark = int(request.form.get(f'm{j}', '0'))
            if subject:
                subjects.append(subject)
                marks.append(mark)
        
        total = sum(marks)
        scale = 25 * len(marks)
        avg = total / scale * 100 if scale > 0 else 0
        g = grade(marks)
        
        df.loc[i] = [exam, semester, branch, year, usn, name, json.dumps(subjects), json.dumps(marks), total, avg, g]
        save_df(df)
        
        return redirect('/faculty')
    
    # GET request - show edit form
    s = df.iloc[i]
    try:
        subjects = json.loads(s['subjects'])
        marks = json.loads(s['marks'])
    except:
        subjects = []
        marks = []
    
    return render_template("edit.html", s=s, subjects=subjects, marks=marks, exams=EXAMS, semesters=SEMESTERS, branches=BRANCHES, years=YEARS)

if __name__ == '__main__':
    app.run(debug=True)