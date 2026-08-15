import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import traceback
import ast
import re
import win32com.client as wincom

# MATPLOTLIB
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ============================================================
# VOICE
# ============================================================

try:
    speak = wincom.Dispatch("SAPI.SpVoice")
except Exception:
    speak = None

voice_enabled = True


def speak_text(text):

    if not voice_enabled:
        return

    if speak is None:
        return

    try:
        speak.Speak(str(text))
    except Exception:
        pass


# ============================================================
# DATABASE
# ============================================================

connection = sqlite3.connect("coding_history.db")
cursor = connection.cursor()


# ============================================================
# USERS TABLE
# ============================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")


# ============================================================
# ERRORS TABLE
# ============================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS errors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    error_type TEXT,
    error_message TEXT,
    line_number INTEGER,
    code TEXT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")


# ============================================================
# SUBMISSIONS TABLE
# ============================================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS submissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    status TEXT,
    code TEXT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")


connection.commit()


# ============================================================
# GLOBAL USER
# ============================================================

current_username = ""


# ============================================================
# SAVE SUBMISSION
# ============================================================

def save_submission(username, status, code):

    cursor.execute("""
    INSERT INTO submissions
    (username, status, code)
    VALUES (?, ?, ?)
    """, (
        username,
        status,
        code
    ))

    connection.commit()


# ============================================================
# SAVE ERROR
# ============================================================

def save_error(
        username,
        error_type,
        error_message,
        line_number,
        code):

    cursor.execute("""
    INSERT INTO errors
    (username, error_type, error_message, line_number, code)
    VALUES (?, ?, ?, ?, ?)
    """, (
        username,
        error_type,
        error_message,
        line_number,
        code
    ))

    connection.commit()


# ============================================================
# ERROR EXPLANATION
# ============================================================

def explain_error(error_type):

    explanations = {

        "ZeroDivisionError":
        "You are trying to divide by zero. Check the denominator.",

        "TypeError":
        "You are using incompatible data types. Check your variable types.",

        "IndexError":
        "You are trying to access a list index that does not exist.",

        "NameError":
        "Python cannot find the variable or function. Check its spelling.",

        "KeyError":
        "The dictionary key you are trying to access does not exist.",

        "AttributeError":
        "The object does not have the attribute or method you are using.",

        "ValueError":
        "The value provided is not suitable for this operation.",

        "SyntaxError":
        "There is a syntax problem. Check brackets, colons and indentation.",

        "IndentationError":
        "The indentation of your Python code is incorrect."
    }

    return explanations.get(
        error_type,
        "An error occurred. Review the error message carefully."
    )


# ============================================================
# CODE QUALITY
# ============================================================

def calculate_quality(code):

    lines = code.splitlines()

    score = 100

    issues = []

    # --------------------------------------------------------
    # LONG LINES
    # --------------------------------------------------------

    for number, line in enumerate(lines, start=1):

        if len(line) > 80:

            issues.append(
                f"Line {number}: Line longer than 80 characters."
            )

            score -= 5

    # --------------------------------------------------------
    # POOR VARIABLE NAMES
    # --------------------------------------------------------

    poor_names = {
        "x",
        "y",
        "z",
        "a",
        "b",
        "c",
        "q"
    }

    for number, line in enumerate(lines, start=1):

        stripped = line.strip()

        if "=" in stripped:

            variable = stripped.split("=")[0].strip()

            # Ignore comparisons such as ==
            if "=" in variable:
                continue

            if variable in poor_names:

                issues.append(
                    f"Line {number}: "
                    f"Variable '{variable}' has a vague name."
                )

                score -= 5

    # --------------------------------------------------------
    # TODO / FIXME
    # --------------------------------------------------------

    for number, line in enumerate(lines, start=1):

        if (
            "TODO" in line.upper()
            or "FIXME" in line.upper()
        ):

            issues.append(
                f"Line {number}: TODO/FIXME found."
            )

            score -= 5

    # --------------------------------------------------------
    # TOO MANY PRINTS
    # --------------------------------------------------------

    print_count = 0

    for line in lines:

        if line.strip().startswith("print("):

            print_count += 1

    if print_count > 5:

        issues.append(
            f"Code contains {print_count} print statements."
        )

        score -= 5

    # --------------------------------------------------------
    # DEEP NESTING
    # --------------------------------------------------------

    for number, line in enumerate(lines, start=1):

        if not line.strip():
            continue

        spaces = len(line) - len(line.lstrip())

        nesting = spaces // 4

        if nesting >= 4:

            issues.append(
                f"Line {number}: Deep nesting detected."
            )

            score -= 5

            break

    score = max(score, 0)

    return score, issues


# ============================================================
# CODE COMPLEXITY
# ============================================================

def calculate_complexity(code):

    lines = code.splitlines()

    actual_lines = [
        line
        for line in lines
        if line.strip()
    ]

    loops = 0
    conditions = 0
    functions = 0
    max_nesting = 0

    # --------------------------------------------------------
    # COUNT LOOPS / CONDITIONS / FUNCTIONS
    # --------------------------------------------------------

    for line in actual_lines:

        stripped = line.strip()

        if (
            stripped.startswith("for ")
            or stripped.startswith("while ")
        ):

            loops += 1

        if (
            stripped.startswith("if ")
            or stripped.startswith("elif ")
            or stripped.startswith("else:")
        ):

            conditions += 1

        if stripped.startswith("def "):

            functions += 1

    # --------------------------------------------------------
    # MAX NESTING
    # --------------------------------------------------------

    for line in lines:

        if not line.strip():
            continue

        spaces = len(line) - len(line.lstrip())

        nesting = spaces // 4

        max_nesting = max(
            max_nesting,
            nesting
        )

    # --------------------------------------------------------
    # COMPLEXITY POINTS
    # --------------------------------------------------------

    points = (
        loops * 2
        + conditions * 2
        + functions
        + max_nesting * 2
    )

    # --------------------------------------------------------
    # COMPLEXITY LEVEL
    # --------------------------------------------------------

    if points <= 5:

        level = "LOW"

    elif points <= 12:

        level = "MEDIUM"

    else:

        level = "HIGH"

    return (
        level,
        points,
        len(actual_lines),
        loops,
        conditions,
        functions,
        max_nesting
    )


# ============================================================
# CLEAR RESULT
# ============================================================

def clear_result():

    result_text.config(
        state="normal"
    )

    result_text.delete(
        "1.0",
        tk.END
    )

    result_text.config(
        state="disabled"
    )


# ============================================================
# SHOW RESULT
# ============================================================

def show_result(text):

    result_text.config(
        state="normal"
    )

    result_text.delete(
        "1.0",
        tk.END
    )

    result_text.insert(
        tk.END,
        text
    )

    result_text.config(
        state="disabled"
    )


# ============================================================
# ANALYZE CODE
# ============================================================

def analyze_code():

    global current_username

    code = code_editor.get(
        "1.0",
        tk.END
    ).strip()

    if not code:

        messagebox.showwarning(
            "No Code",
            "Please enter some Python code first."
        )

        return

    clear_result()

    # ========================================================
    # SYNTAX CHECK
    # ========================================================

    try:

        compile(
            code,
            "<student_code>",
            "exec"
        )

    except SyntaxError as e:

        error_type = "SyntaxError"

        error_message = e.msg

        line_number = e.lineno

        save_error(
            current_username,
            error_type,
            error_message,
            line_number,
            code
        )

        save_submission(
            current_username,
            "FAILED",
            code
        )

        result = f"""
❌ SYNTAX ERROR

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Error Type:
{error_type}

Line:
{line_number}

Problem:
{error_message}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 MENTOR

{explain_error(error_type)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🛠️ POSSIBLE FIX

Check your brackets, colons,
indentation and Python syntax.

"""

        show_result(result)

        speak_text(
            f"Syntax error detected on line "
            f"{line_number}. "
            f"{explain_error(error_type)}"
        )

        update_dashboard()

        return

    # ========================================================
    # EXECUTION
    # ========================================================

    try:

        exec(
            code,
            {
                "__builtins__": __builtins__
            }
        )

    except Exception as e:

        error_type = type(e).__name__

        error_message = str(e)

        line_number = None

        try:

            tb = traceback.extract_tb(
                e.__traceback__
            )

            if tb:

                line_number = tb[-1].lineno

        except Exception:

            pass

        save_error(
            current_username,
            error_type,
            error_message,
            line_number,
            code
        )

        save_submission(
            current_username,
            "FAILED",
            code
        )

        result = f"""
❌ RUNTIME ERROR

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Error Type:
{error_type}

Problem:
{error_message}

Line:
{line_number}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🔎 ERROR LOCATION

The error occurred around line {line_number}.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

💡 MENTOR

{explain_error(error_type)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🛠️ POSSIBLE FIX

Review the affected line and check
the variables and data types being used.

"""

        show_result(result)

        speak_text(
            f"Runtime error detected. "
            f"The error is {error_type}. "
            f"{explain_error(error_type)}"
        )

        update_dashboard()

        return

    # ========================================================
    # QUALITY
    # ========================================================

    score, issues = calculate_quality(
        code
    )

    # ========================================================
    # COMPLEXITY
    # ========================================================

    (
        complexity,
        points,
        lines,
        loops,
        conditions,
        functions,
        nesting
    ) = calculate_complexity(code)

    # ========================================================
    # QUALITY RECOMMENDATION
    # ========================================================

    if score >= 90:

        recommendation = (
            "Excellent code quality! "
            "Your code is clean and readable."
        )

    elif score >= 75:

        recommendation = (
            "Good code. A few improvements "
            "can make it cleaner."
        )

    elif score >= 50:

        recommendation = (
            "Your code works, but it needs "
            "some cleanup."
        )

    else:

        recommendation = (
            "Your code needs significant improvement."
        )

    # ========================================================
    # COMPLEXITY RECOMMENDATION
    # ========================================================

    if complexity == "LOW":

        complexity_tip = (
            "Your code is simple and easy to understand."
        )

    elif complexity == "MEDIUM":

        complexity_tip = (
            "Your code has moderate complexity. "
            "Consider breaking large sections "
            "into smaller functions."
        )

    else:

        complexity_tip = (
            "Your code is highly complex. "
            "Try reducing nesting and splitting "
            "the program into smaller functions."
        )

    # ========================================================
    # AUTOMATIC IMPROVEMENTS
    # ========================================================

    improvements = []

    try:

        tree = ast.parse(code)

        poor_names = {
            "x",
            "y",
            "z",
            "a",
            "b",
            "c"
        }

        # ----------------------------------------------------
        # VARIABLE NAMES
        # ----------------------------------------------------

        found_names = set()

        for node in ast.walk(tree):

            if isinstance(node, ast.Name):

                if node.id in poor_names:

                    found_names.add(
                        node.id
                    )

        for name in sorted(found_names):

            improvements.append(
                f"Consider replacing '{name}' "
                "with a descriptive variable name."
            )

        # ----------------------------------------------------
        # FUNCTIONS
        # ----------------------------------------------------

        if functions == 0 and lines > 10:

            improvements.append(
                "Consider using functions to divide "
                "your program into reusable components."
            )

        # ----------------------------------------------------
        # INPUT VALIDATION
        # ----------------------------------------------------

        if "input(" in code:

            if "try:" not in code:

                improvements.append(
                    "Consider using try/except "
                    "for input validation."
                )

        # ----------------------------------------------------
        # MAGIC NUMBERS
        # ----------------------------------------------------

        number_pattern = (
            r"(?<![\w.])\d{2,}(?![\w.])"
        )

        if re.findall(
            number_pattern,
            code
        ):

            improvements.append(
                "Consider storing frequently used "
                "numbers in named constants."
            )

        # ----------------------------------------------------
        # COMMENTS
        # ----------------------------------------------------

        comment_count = 0

        for line in code.splitlines():

            if line.strip().startswith("#"):

                comment_count += 1

        if lines > 10 and comment_count == 0:

            improvements.append(
                "Consider adding comments to explain "
                "important sections of your code."
            )

    except Exception:

        pass

    # ========================================================
    # SAVE SUCCESSFUL SUBMISSION
    # ========================================================

    save_submission(
        current_username,
        "SUCCESS",
        code
    )

    # ========================================================
    # BUILD RESULT
    # ========================================================

    result = f"""
✅ CODE EXECUTED SUCCESSFULLY

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🧹 CODE QUALITY

Score:
{score}/100

"""

    # --------------------------------------------------------
    # QUALITY ISSUES
    # --------------------------------------------------------

    if issues:

        result += "⚠️ Issues Found:\n\n"

        for issue in issues:

            result += f"• {issue}\n"

    else:

        result += (
            "🎉 No major quality issues detected.\n"
        )

    # --------------------------------------------------------
    # COMPLEXITY
    # --------------------------------------------------------

    result += f"""

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📐 CODE COMPLEXITY

Level:
{complexity}

Complexity Points:
{points}

Lines:
{lines}

Loops:
{loops}

Conditions:
{conditions}

Functions:
{functions}

Maximum Nesting:
{nesting}

💡 Complexity Mentor:

{complexity_tip}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🧑‍💻 MENTOR JUDGEMENT

{recommendation}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✨ AUTOMATIC IMPROVEMENTS

"""

    # --------------------------------------------------------
    # IMPROVEMENTS
    # --------------------------------------------------------

    if improvements:

        for number, improvement in enumerate(
            improvements,
            start=1
        ):

            result += (
                f"{number}. {improvement}\n"
            )

    else:

        result += (
            "🎉 Your code already follows "
            "several good practices!"
        )

    # ========================================================
    # SHOW RESULT
    # ========================================================

    show_result(
        result
    )

    # ========================================================
    # VOICE MENTOR
    # ========================================================

    speak_text(
        f"Code executed successfully. "
        f"Your quality score is "
        f"{score} out of 100. "
        f"Your complexity level is "
        f"{complexity}. "
        f"Mentor judgement: "
        f"{recommendation}"
    )

    update_dashboard()


# ============================================================
# CLEAR CODE
# ============================================================

def clear_code():

    code_editor.delete(
        "1.0",
        tk.END
    )

    clear_result()


# ============================================================
# UPDATE DASHBOARD
# ============================================================

def update_dashboard():

    cursor.execute("""
    SELECT status
    FROM submissions
    WHERE username = ?
    """, (
        current_username,
    ))

    submissions = cursor.fetchall()

    total = len(submissions)

    successful = sum(
        1
        for row in submissions
        if row[0] == "SUCCESS"
    )

    failed = sum(
        1
        for row in submissions
        if row[0] == "FAILED"
    )

    if total:

        rate = (
            successful / total
        ) * 100

    else:

        rate = 0

    total_label.config(
        text=str(total)
    )

    success_label.config(
        text=str(successful)
    )

    failed_label.config(
        text=str(failed)
    )

    rate_label.config(
        text=f"{rate:.1f}%"
    )


# ============================================================
# SHOW HISTORY
# ============================================================

def show_history():

    cursor.execute("""
    SELECT
        error_type,
        error_message,
        line_number,
        timestamp
    FROM errors
    WHERE username = ?
    ORDER BY id DESC
    """, (
        current_username,
    ))

    records = cursor.fetchall()

    clear_result()

    if not records:

        show_result(
            "🎉 No errors recorded yet!"
        )

        speak_text(
            "No errors recorded yet."
        )

        return

    text = (
        f"📜 {current_username.upper()}'S ERROR HISTORY\n"
        "\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    )

    for record in records:

        text += (
            f"❌ Error: {record[0]}\n"
            f"Message: {record[1]}\n"
            f"Line: {record[2]}\n"
            f"Time: {record[3]}\n"
            "\n"
            "────────────────────────────\n\n"
        )

    show_result(
        text
    )


# ============================================================
# BUG PATTERNS
# ============================================================

def show_bug_patterns():

    cursor.execute("""
    SELECT error_type
    FROM errors
    WHERE username = ?
    """, (
        current_username,
    ))

    records = cursor.fetchall()

    if not records:

        show_result(
            "🎉 No bug data available yet!"
        )

        speak_text(
            "No bug data available yet."
        )

        return

    counts = {}

    for record in records:

        error = record[0]

        counts[error] = (
            counts.get(error, 0) + 1
        )

    sorted_errors = sorted(
        counts.items(),
        key=lambda x: x[1],
        reverse=True
    )

    text = (
        f"🐛 {current_username.upper()}'S BUG PATTERNS\n"
        "\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    )

    for error, count in sorted_errors:

        text += (
            f"❌ {error}: {count} time(s)\n"
        )

    most_common = sorted_errors[0][0]

    text += (
        "\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"⚠️ MOST COMMON BUG\n\n"
        f"{most_common}\n\n"
    )

    recommendations = {

        "TypeError":
        "Practice data types and type conversion.",

        "IndexError":
        "Practice lists, indexing and slicing.",

        "NameError":
        "Practice variables and scope.",

        "KeyError":
        "Practice dictionaries.",

        "ValueError":
        "Practice type conversion and validation.",

        "ZeroDivisionError":
        "Practice conditions and arithmetic.",

        "AttributeError":
        "Practice objects and methods.",

        "SyntaxError":
        "Practice Python syntax and indentation."
    }

    recommendation = recommendations.get(
        most_common,
        "Review your error history and keep practicing."
    )

    text += (
        f"📚 RECOMMENDATION\n\n"
        f"{recommendation}"
    )

    show_result(
        text
    )

    speak_text(
        f"Your most common error is "
        f"{most_common}. "
        f"Recommendation: "
        f"{recommendation}"
    )


# ============================================================
# GRAPHS & CHARTS
# ============================================================

def show_graphs():

    # ========================================================
    # GET DATA
    # ========================================================

    cursor.execute("""
    SELECT status
    FROM submissions
    WHERE username = ?
    """, (
        current_username,
    ))

    records = cursor.fetchall()

    successful = sum(
        1
        for row in records
        if row[0] == "SUCCESS"
    )

    failed = sum(
        1
        for row in records
        if row[0] == "FAILED"
    )

    total = successful + failed

    if total > 0:

        success_rate = (
            successful / total
        ) * 100

    else:

        success_rate = 0

    # ========================================================
    # GRAPH WINDOW
    # ========================================================

    graph_window = tk.Toplevel(
        root
    )

    graph_window.title(
        f"{current_username}'s Coding Analytics"
    )

    graph_window.geometry(
        "1050x800"
    )

    graph_window.minsize(
        900,
        650
    )

    graph_window.configure(
        bg="#101820"
    )

    # ========================================================
    # TITLE
    # ========================================================

    title = tk.Label(
        graph_window,
        text="📊 CODING ANALYTICS",
        font=("Segoe UI", 24, "bold"),
        bg="#101820",
        fg="white"
    )

    title.pack(
        pady=(20, 5)
    )

    subtitle = tk.Label(
        graph_window,
        text=f"Performance report for {current_username}",
        font=("Segoe UI", 11),
        bg="#101820",
        fg="#aaaaaa"
    )

    subtitle.pack(
        pady=(0, 10)
    )

    # ========================================================
    # STATISTICS
    # ========================================================

    stats = tk.Label(
        graph_window,
        text=(
            f"Successful Programs: {successful}     "
            f"Failed Programs: {failed}     "
            f"Total Programs: {total}"
        ),
        font=("Segoe UI", 12),
        bg="#101820",
        fg="#dddddd"
    )

    stats.pack(
        pady=(0, 10)
    )

    # ========================================================
    # MATPLOTLIB FIGURE
    # ========================================================

    figure = Figure(
        figsize=(9, 6),
        dpi=100
    )

    # ========================================================
    # PIE CHART
    # ========================================================

    pie_ax = figure.add_subplot(
        221
    )

    if total > 0:

        pie_ax.pie(
            [successful, failed],
            labels=[
                "Successful",
                "Failed"
            ],
            autopct="%1.1f%%",
            startangle=90
        )

    else:

        pie_ax.text(
            0.5,
            0.5,
            "No data",
            ha="center",
            va="center"
        )

    pie_ax.set_title(
        "Success vs Failure"
    )

    # ========================================================
    # BAR CHART
    # ========================================================

    bar_ax = figure.add_subplot(
        222
    )

    bar_ax.bar(
        [
            "Successful",
            "Failed"
        ],
        [
            successful,
            failed
        ]
    )

    bar_ax.set_title(
        "Program Results"
    )

    bar_ax.set_ylabel(
        "Number of Programs"
    )

    # ========================================================
    # SUCCESS RATE
    # ========================================================

    rate_ax = figure.add_subplot(
        223
    )

    rate_ax.bar(
        [
            "Success Rate"
        ],
        [
            success_rate
        ]
    )

    rate_ax.set_ylim(
        0,
        100
    )

    rate_ax.set_ylabel(
        "Percentage"
    )

    rate_ax.set_title(
        "Coding Success Rate"
    )

    rate_ax.text(
        0,
        success_rate + 3,
        f"{success_rate:.1f}%",
        ha="center",
        fontweight="bold"
    )

    # ========================================================
    # SUMMARY PANEL
    # ========================================================

    summary_ax = figure.add_subplot(
        224
    )

    summary_ax.axis(
        "off"
    )

    summary_text = (
        "CODING SUMMARY\n\n"
        f"Total Programs\n"
        f"{total}\n\n"
        f"Successful\n"
        f"{successful}\n\n"
        f"Failed\n"
        f"{failed}\n\n"
        f"Success Rate\n"
        f"{success_rate:.1f}%"
    )

    summary_ax.text(
        0.5,
        0.5,
        summary_text,
        ha="center",
        va="center",
        fontsize=12
    )

    # ========================================================
    # LAYOUT
    # ========================================================

    figure.tight_layout(
        pad=3
    )

    # ========================================================
    # EMBED MATPLOTLIB INTO TKINTER
    # ========================================================

    canvas = FigureCanvasTkAgg(
        figure,
        master=graph_window
    )

    canvas.draw()

    canvas_widget = canvas.get_tk_widget()

    canvas_widget.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=5
    )

    # ========================================================
    # CLOSE BUTTON
    # ========================================================

    close_button = tk.Button(
        graph_window,
        text="✖ CLOSE",
        font=("Segoe UI", 11, "bold"),
        bg="#24333d",
        fg="white",
        relief="flat",
        padx=25,
        pady=8,
        command=graph_window.destroy
    )

    close_button.pack(
        pady=10
    )

    # ========================================================
    # VOICE
    # ========================================================

    speak_text(
        f"Here are your coding analytics. "
        f"You completed {total} programs. "
        f"{successful} were successful and "
        f"{failed} failed. "
        f"Your success rate is "
        f"{success_rate:.1f} percent."
    )


# ============================================================
# LOGIN
# ============================================================

def login():

    global current_username

    username = username_entry.get().strip()

    if not username:

        messagebox.showwarning(
            "Username Required",
            "Please enter your username."
        )

        return

    cursor.execute("""
    SELECT username
    FROM users
    WHERE username = ?
    """, (
        username,
    ))

    user = cursor.fetchone()

    if user is None:

        cursor.execute("""
        INSERT INTO users (username)
        VALUES (?)
        """, (
            username,
        ))

        connection.commit()

        welcome = (
            f"Welcome to AI Code Mentor, "
            f"{username}!"
        )

    else:

        welcome = (
            f"Welcome back, "
            f"{username}!"
        )

    current_username = username

    login_frame.pack_forget()

    main_frame.pack(
        fill="both",
        expand=True
    )

    user_label.config(
        text=f"👤 {current_username}"
    )

    update_dashboard()

    speak_text(
        welcome
    )


# ============================================================
# VOICE TOGGLE
# ============================================================

def toggle_voice():

    global voice_enabled

    voice_enabled = not voice_enabled

    if voice_enabled:

        voice_button.config(
            text="🔊 Voice ON"
        )

        speak_text(
            "Voice mentor enabled."
        )

    else:

        voice_button.config(
            text="🔇 Voice OFF"
        )


# ============================================================
# MAIN WINDOW
# ============================================================

root = tk.Tk()

root.title(
    "🤖 AI Code Mentor"
)

root.geometry(
    "1200x750"
)

root.minsize(
    1000,
    650
)

root.configure(
    bg="#101820"
)


# ============================================================
# STYLE
# ============================================================

style = ttk.Style()

try:

    style.theme_use(
        "clam"
    )

except Exception:

    pass


style.configure(
    "TButton",
    font=("Segoe UI", 11),
    padding=8
)

style.configure(
    "TLabel",
    background="#101820",
    foreground="white",
    font=("Segoe UI", 11)
)


# ============================================================
# LOGIN FRAME
# ============================================================

login_frame = tk.Frame(
    root,
    bg="#101820"
)

login_frame.pack(
    fill="both",
    expand=True
)


login_title = tk.Label(
    login_frame,
    text="🤖 AI CODE MENTOR",
    font=("Segoe UI", 30, "bold"),
    bg="#101820",
    fg="white"
)

login_title.pack(
    pady=(130, 15)
)


login_subtitle = tk.Label(
    login_frame,
    text="Your Personal Python Coding Mentor",
    font=("Segoe UI", 14),
    bg="#101820",
    fg="#aaaaaa"
)

login_subtitle.pack(
    pady=5
)


username_entry = tk.Entry(
    login_frame,
    font=("Segoe UI", 14),
    width=30,
    justify="center"
)

username_entry.pack(
    pady=25,
    ipady=8
)


login_button = tk.Button(
    login_frame,
    text="🚀 ENTER MENTOR",
    font=("Segoe UI", 13, "bold"),
    bg="#2d89ef",
    fg="white",
    padx=30,
    pady=10,
    relief="flat",
    command=login
)

login_button.pack(
    pady=10
)


# ============================================================
# MAIN FRAME
# ============================================================

main_frame = tk.Frame(
    root,
    bg="#101820"
)


# ============================================================
# HEADER
# ============================================================

header = tk.Frame(
    main_frame,
    bg="#17232c",
    height=70
)

header.pack(
    fill="x"
)

header.pack_propagate(
    False
)


title_label = tk.Label(
    header,
    text="🤖 AI CODE MENTOR",
    font=("Segoe UI", 20, "bold"),
    bg="#17232c",
    fg="white"
)

title_label.pack(
    side="left",
    padx=25
)


user_label = tk.Label(
    header,
    text="👤 User",
    font=("Segoe UI", 12),
    bg="#17232c",
    fg="#cccccc"
)

user_label.pack(
    side="right",
    padx=25
)


# ============================================================
# CONTENT
# ============================================================

content = tk.Frame(
    main_frame,
    bg="#101820"
)

content.pack(
    fill="both",
    expand=True,
    padx=15,
    pady=15
)


# ============================================================
# SIDEBAR
# ============================================================

sidebar = tk.Frame(
    content,
    bg="#17232c",
    width=190
)

sidebar.pack(
    side="left",
    fill="y",
    padx=(0, 15)
)

sidebar.pack_propagate(
    False
)


sidebar_title = tk.Label(
    sidebar,
    text="MENU",
    font=("Segoe UI", 12, "bold"),
    bg="#17232c",
    fg="#aaaaaa"
)

sidebar_title.pack(
    pady=(20, 15)
)


# ============================================================
# ANALYZE BUTTON
# ============================================================

analyze_button = tk.Button(
    sidebar,
    text="💻 Analyze Code",
    font=("Segoe UI", 11),
    bg="#2d89ef",
    fg="white",
    relief="flat",
    command=lambda: code_editor.focus()
)

analyze_button.pack(
    fill="x",
    padx=12,
    pady=5
)


# ============================================================
# HISTORY BUTTON
# ============================================================

history_button = tk.Button(
    sidebar,
    text="📜 Error History",
    font=("Segoe UI", 11),
    bg="#24333d",
    fg="white",
    relief="flat",
    command=show_history
)

history_button.pack(
    fill="x",
    padx=12,
    pady=5
)


# ============================================================
# BUG PATTERNS BUTTON
# ============================================================

bugs_button = tk.Button(
    sidebar,
    text="🐛 Bug Patterns",
    font=("Segoe UI", 11),
    bg="#24333d",
    fg="white",
    relief="flat",
    command=show_bug_patterns
)

bugs_button.pack(
    fill="x",
    padx=12,
    pady=5
)


# ============================================================
# GRAPHS BUTTON
# ============================================================

graphs_button = tk.Button(
    sidebar,
    text="📊 Graphs & Charts",
    font=("Segoe UI", 11),
    bg="#24333d",
    fg="white",
    relief="flat",
    command=show_graphs
)

graphs_button.pack(
    fill="x",
    padx=12,
    pady=5
)


# ============================================================
# VOICE BUTTON
# ============================================================

voice_button = tk.Button(
    sidebar,
    text="🔊 Voice ON",
    font=("Segoe UI", 11),
    bg="#24333d",
    fg="white",
    relief="flat",
    command=toggle_voice
)

voice_button.pack(
    fill="x",
    padx=12,
    pady=5
)


# ============================================================
# RIGHT AREA
# ============================================================

right_area = tk.Frame(
    content,
    bg="#101820"
)

right_area.pack(
    side="left",
    fill="both",
    expand=True
)


# ============================================================
# DASHBOARD CARDS
# ============================================================

stats_frame = tk.Frame(
    right_area,
    bg="#101820"
)

stats_frame.pack(
    fill="x",
    pady=(0, 12)
)


def create_card(parent, title):

    frame = tk.Frame(
        parent,
        bg="#17232c",
        width=160,
        height=75
    )

    frame.pack(
        side="left",
        padx=(0, 10)
    )

    frame.pack_propagate(
        False
    )

    title_label = tk.Label(
        frame,
        text=title,
        font=("Segoe UI", 9),
        bg="#17232c",
        fg="#aaaaaa"
    )

    title_label.pack(
        pady=(8, 0)
    )

    value_label = tk.Label(
        frame,
        text="0",
        font=("Segoe UI", 18, "bold"),
        bg="#17232c",
        fg="white"
    )

    value_label.pack()

    return value_label


total_label = create_card(
    stats_frame,
    "TOTAL PROGRAMS"
)

success_label = create_card(
    stats_frame,
    "SUCCESSFUL"
)

failed_label = create_card(
    stats_frame,
    "FAILED"
)

rate_label = create_card(
    stats_frame,
    "SUCCESS RATE"
)


# ============================================================
# EDITOR LABEL
# ============================================================

editor_label = tk.Label(
    right_area,
    text="💻 Python Code",
    font=("Segoe UI", 13, "bold"),
    bg="#101820",
    fg="white"
)

editor_label.pack(
    anchor="w"
)


# ============================================================
# CODE EDITOR
# ============================================================

editor_frame = tk.Frame(
    right_area,
    bg="#17232c"
)

editor_frame.pack(
    fill="both",
    expand=True,
    pady=(5, 10)
)


code_editor = tk.Text(
    editor_frame,
    bg="#0b1115",
    fg="#eeeeee",
    insertbackground="white",
    font=("Consolas", 12),
    undo=True,
    wrap="none",
    padx=15,
    pady=15
)

code_editor.pack(
    fill="both",
    expand=True
)


# ============================================================
# BUTTON FRAME
# ============================================================

button_frame = tk.Frame(
    right_area,
    bg="#101820"
)

button_frame.pack(
    fill="x",
    pady=(0, 10)
)


# ============================================================
# RUN BUTTON
# ============================================================

run_button = tk.Button(
    button_frame,
    text="▶️ RUN & ANALYZE",
    font=("Segoe UI", 11, "bold"),
    bg="#2d89ef",
    fg="white",
    relief="flat",
    padx=20,
    pady=8,
    command=analyze_code
)

run_button.pack(
    side="left",
    padx=(0, 8)
)


# ============================================================
# CLEAR BUTTON
# ============================================================

clear_button = tk.Button(
    button_frame,
    text="🗑️ CLEAR",
    font=("Segoe UI", 11),
    bg="#24333d",
    fg="white",
    relief="flat",
    padx=20,
    pady=8,
    command=clear_code
)

clear_button.pack(
    side="left"
)


# ============================================================
# RESULT LABEL
# ============================================================

result_label = tk.Label(
    right_area,
    text="🔎 Analysis Result",
    font=("Segoe UI", 13, "bold"),
    bg="#101820",
    fg="white"
)

result_label.pack(
    anchor="w"
)


# ============================================================
# RESULT FRAME
# ============================================================

result_frame = tk.Frame(
    right_area,
    bg="#17232c"
)

result_frame.pack(
    fill="both",
    expand=True
)


# ============================================================
# RESULT TEXT
# ============================================================

result_text = tk.Text(
    result_frame,
    bg="#0b1115",
    fg="#eeeeee",
    font=("Consolas", 10),
    wrap="word",
    padx=15,
    pady=15,
    state="disabled"
)

result_text.pack(
    fill="both",
    expand=True
)


# ============================================================
# CLOSE DATABASE SAFELY
# ============================================================

def close_application():

    try:

        connection.commit()
        connection.close()

    except Exception:

        pass

    root.destroy()


# ============================================================
# WINDOW CLOSE
# ============================================================

root.protocol(
    "WM_DELETE_WINDOW",
    close_application
)


# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()