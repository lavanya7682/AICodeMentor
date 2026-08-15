import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import traceback
import ast
import re
import random
import io
import contextlib

# ============================================================
# OPTIONAL MATPLOTLIB
# ============================================================

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False


# ============================================================
# VOICE
# ============================================================

try:
    import win32com.client as wincom

    speak = wincom.Dispatch("SAPI.SpVoice")
    voice_available = True

except Exception:
    speak = None
    voice_available = False

voice_enabled = True


def speak_text(text):

    if not voice_enabled or not voice_available:
        return

    try:
        speak.Speak(str(text))
    except Exception:
        pass


# ============================================================
# DATABASE
# ============================================================

connection = sqlite3.connect(
    "coding_history.db"
)

cursor = connection.cursor()


cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")


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
# GLOBALS
# ============================================================

current_username = ""

practice_difficulty = ""
practice_type = ""

current_question = None

practice_score = 0
practice_total = 0
practice_correct = 0
practice_wrong = 0
practice_answered = False


# ============================================================
# PRACTICE QUESTION DATABASE
# ============================================================

PRACTICE_DATA = {

    "Easy": {

        "MCQ": [

            {
                "question":
                "What is the output of: print(2 + 3 * 2)?",

                "options": [
                    "10",
                    "8",
                    "12",
                    "7"
                ],

                "answer":
                "8",

                "explanation":
                "Multiplication happens before addition. "
                "3 multiplied by 2 is 6 and 2 plus 6 is 8."
            },

            {
                "question":
                "Which keyword is used to define a function?",

                "options": [
                    "function",
                    "define",
                    "def",
                    "fun"
                ],

                "answer":
                "def",

                "explanation":
                "Python uses the def keyword to create a function."
            },

            {
                "question":
                "Which data type stores True or False?",

                "options": [
                    "int",
                    "str",
                    "bool",
                    "float"
                ],

                "answer":
                "bool",

                "explanation":
                "Boolean values in Python are True and False."
            },

            {
                "question":
                "What does len([10, 20, 30]) return?",

                "options": [
                    "2",
                    "3",
                    "4",
                    "30"
                ],

                "answer":
                "3",

                "explanation":
                "The list contains three elements."
            }
        ],

        "Code Writing": [

            {
                "question":
                "Write Python code to print 'Hello World'.",

                "answer":
                "print('Hello World')",

                "explanation":
                "Use Python's print function to display Hello World."
            },

            {
                "question":
                "Write Python code to calculate the sum of two "
                "numbers stored in variables a and b.",

                "answer":
                "print(a + b)",

                "explanation":
                "The plus operator adds the values stored in a and b."
            }
        ],

        "Code Output": [

            {
                "question":
                "What is the output?\n\n"
                "x = 5\n"
                "y = 2\n"
                "print(x + y)",

                "answer":
                "7",

                "explanation":
                "5 plus 2 equals 7."
            },

            {
                "question":
                "What is the output?\n\n"
                "name = 'Python'\n"
                "print(len(name))",

                "answer":
                "6",

                "explanation":
                "The word Python contains six characters."
            }
        ]
    },


    "Medium": {

        "MCQ": [

            {
                "question":
                "What is the output of:\n"
                "numbers = [1, 2, 3]\n"
                "print(numbers[-1])",

                "options": [
                    "1",
                    "2",
                    "3",
                    "Error"
                ],

                "answer":
                "3",

                "explanation":
                "Index -1 refers to the last element of a list."
            },

            {
                "question":
                "Which method adds an element to the end of a list?",

                "options": [
                    "add()",
                    "append()",
                    "insert_end()",
                    "push()"
                ],

                "answer":
                "append()",

                "explanation":
                "The append method adds an element to the end of a list."
            },

            {
                "question":
                "What does range(5) generate?",

                "options": [
                    "1,2,3,4,5",
                    "0,1,2,3,4",
                    "0,1,2,3,4,5",
                    "5 only"
                ],

                "answer":
                "0,1,2,3,4",

                "explanation":
                "range(5) starts at zero and stops before five."
            }
        ],

        "Code Writing": [

            {
                "question":
                "Write a Python program that prints numbers "
                "from 1 to 5 using a for loop.",

                "answer":
                "for i in range(1, 6):\n"
                "    print(i)",

                "explanation":
                "range(1, 6) generates the numbers 1 through 5."
            },

            {
                "question":
                "Write a Python program to check whether a number "
                "is even or odd.",

                "answer":
                "if n % 2 == 0:\n"
                "    print('Even')\n"
                "else:\n"
                "    print('Odd')",

                "explanation":
                "A number is even when its remainder after division "
                "by 2 is zero."
            }
        ],

        "Code Output": [

            {
                "question":
                "What is the output?\n\n"
                "total = 0\n"
                "for i in range(1, 4):\n"
                "    total += i\n"
                "print(total)",

                "answer":
                "6",

                "explanation":
                "The loop calculates 1 plus 2 plus 3, which equals 6."
            },

            {
                "question":
                "What is the output?\n\n"
                "numbers = [10, 20, 30]\n"
                "numbers.append(40)\n"
                "print(len(numbers))",

                "answer":
                "4",

                "explanation":
                "append adds 40, so the list contains four elements."
            }
        ]
    },


    "Tough": {

        "MCQ": [

            {
                "question":
                "What is the output?\n\n"
                "x = [1, 2, 3]\n"
                "y = x\n"
                "y.append(4)\n"
                "print(x)",

                "options": [
                    "[1, 2, 3]",
                    "[1, 2, 3, 4]",
                    "[4]",
                    "Error"
                ],

                "answer":
                "[1, 2, 3, 4]",

                "explanation":
                "y refers to the same list object as x. "
                "Therefore appending through y also changes x."
            },

            {
                "question":
                "Which concept allows a function to call itself?",

                "options": [
                    "Iteration",
                    "Inheritance",
                    "Recursion",
                    "Encapsulation"
                ],

                "answer":
                "Recursion",

                "explanation":
                "Recursion occurs when a function calls itself."
            },

            {
                "question":
                "What is the average-case lookup complexity "
                "of a Python dictionary?",

                "options": [
                    "O(n)",
                    "O(log n)",
                    "O(1)",
                    "O(n²)"
                ],

                "answer":
                "O(1)",

                "explanation":
                "Python dictionaries use hash tables, providing "
                "average constant-time lookup."
            }
        ],

        "Code Writing": [

            {
                "question":
                "Write Python code to find the largest number "
                "in a list WITHOUT using max().",

                "answer":
                "largest = numbers[0]\n"
                "for number in numbers:\n"
                "    if number > largest:\n"
                "        largest = number\n"
                "print(largest)",

                "explanation":
                "Start with the first element and update largest "
                "whenever a larger value is found."
            },

            {
                "question":
                "Write Python code to count the frequency of each "
                "element in a list using a dictionary.",

                "answer":
                "frequency = {}\n"
                "for item in numbers:\n"
                "    frequency[item] = frequency.get(item, 0) + 1",

                "explanation":
                "The dictionary stores each item and increments "
                "its frequency whenever the item appears."
            }
        ],

        "Code Output": [

            {
                "question":
                "What is the output?\n\n"
                "def fun(x):\n"
                "    if x == 0:\n"
                "        return 1\n"
                "    return x * fun(x - 1)\n\n"
                "print(fun(4))",

                "answer":
                "24",

                "explanation":
                "The recursive function calculates "
                "4 multiplied by 3 multiplied by 2 multiplied by 1."
            },

            {
                "question":
                "What is the output?\n\n"
                "data = {'a': 1, 'b': 2}\n"
                "data['c'] = data['a'] + data['b']\n"
                "print(data['c'])",

                "answer":
                "3",

                "explanation":
                "data a is 1 and data b is 2, so their sum is 3."
            }
        ]
    }
}


# ============================================================
# DATABASE FUNCTIONS
# ============================================================

def save_submission(username, status, code):

    cursor.execute(
        """
        INSERT INTO submissions
        (username, status, code)
        VALUES (?, ?, ?)
        """,
        (
            username,
            status,
            code
        )
    )

    connection.commit()


def save_error(
        username,
        error_type,
        error_message,
        line_number,
        code):

    cursor.execute(
        """
        INSERT INTO errors
        (username, error_type, error_message,
         line_number, code)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            username,
            error_type,
            error_message,
            line_number,
            code
        )
    )

    connection.commit()


# ============================================================
# ERROR EXPLANATION
# ============================================================

def explain_error(error_type):

    explanations = {

        "ZeroDivisionError":
        "You are trying to divide by zero. "
        "Check the denominator.",

        "TypeError":
        "You are using incompatible data types. "
        "Check your variable types.",

        "IndexError":
        "You are trying to access a list index "
        "that does not exist.",

        "NameError":
        "Python cannot find the variable or function. "
        "Check its spelling.",

        "KeyError":
        "The dictionary key you are trying to access "
        "does not exist.",

        "AttributeError":
        "The object does not have the attribute "
        "or method you are using.",

        "ValueError":
        "The value provided is not suitable "
        "for this operation.",

        "SyntaxError":
        "There is a syntax problem. Check brackets, "
        "colons and indentation.",

        "IndentationError":
        "The indentation of your Python code is incorrect."
    }

    return explanations.get(
        error_type,
        "An error occurred. Review the error message carefully."
    )


# ============================================================
# QUALITY ANALYZER
# ============================================================

def calculate_quality(code):

    lines = code.splitlines()

    score = 100

    issues = []

    for number, line in enumerate(
        lines,
        start=1
    ):

        if len(line) > 80:

            issues.append(
                f"Line {number}: "
                "Line longer than 80 characters."
            )

            score -= 5

    poor_names = {
        "x",
        "y",
        "z",
        "a",
        "b",
        "c",
        "q"
    }

    for number, line in enumerate(
        lines,
        start=1
    ):

        stripped = line.strip()

        if "=" in stripped:

            variable = (
                stripped
                .split("=")[0]
                .strip()
            )

            if variable in poor_names:

                issues.append(
                    f"Line {number}: "
                    f"Variable '{variable}' "
                    "has a vague name."
                )

                score -= 5

    for number, line in enumerate(
        lines,
        start=1
    ):

        if (
            "TODO" in line.upper()
            or
            "FIXME" in line.upper()
        ):

            issues.append(
                f"Line {number}: TODO/FIXME found."
            )

            score -= 5

    print_count = sum(
        1
        for line in lines
        if line.strip().startswith("print(")
    )

    if print_count > 5:

        issues.append(
            f"Code contains {print_count} "
            "print statements."
        )

        score -= 5

    for number, line in enumerate(
        lines,
        start=1
    ):

        if not line.strip():
            continue

        spaces = (
            len(line)
            -
            len(line.lstrip())
        )

        nesting = spaces // 4

        if nesting >= 4:

            issues.append(
                f"Line {number}: "
                "Deep nesting detected."
            )

            score -= 5

            break

    score = max(
        0,
        score
    )

    return score, issues


# ============================================================
# COMPLEXITY ANALYZER
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

    for line in actual_lines:

        stripped = line.strip()

        if (
            stripped.startswith("for ")
            or
            stripped.startswith("while ")
        ):

            loops += 1

        if (
            stripped.startswith("if ")
            or
            stripped.startswith("elif ")
            or
            stripped.startswith("else:")
        ):

            conditions += 1

        if stripped.startswith("def "):

            functions += 1

    for line in lines:

        if not line.strip():
            continue

        spaces = (
            len(line)
            -
            len(line.lstrip())
        )

        max_nesting = max(
            max_nesting,
            spaces // 4
        )

    points = (
        loops * 2
        +
        conditions * 2
        +
        functions
        +
        max_nesting * 2
    )

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
# RUN & ANALYZE
# ============================================================

def analyze_code():

    code = code_editor.get(
        "1.0",
        tk.END
    ).strip()

    if not code:

        messagebox.showwarning(
            "No Code",
            "Please enter Python code first."
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

        save_error(
            current_username,
            "SyntaxError",
            e.msg,
            e.lineno,
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
SyntaxError

Line:
{e.lineno}

Problem:
{e.msg}

💡 MENTOR

{explain_error("SyntaxError")}

🛠️ POSSIBLE FIX

Check brackets, quotes, colons
and indentation.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

        show_result(result)

        speak_text(
            f"Syntax error detected on line "
            f"{e.lineno}. "
            f"{explain_error('SyntaxError')}"
        )

        update_dashboard()

        return

    # ========================================================
    # EXECUTION + OUTPUT CAPTURE
    # ========================================================

    output_buffer = io.StringIO()

    try:

        execution_namespace = {
            "__builtins__": __builtins__
        }

        with contextlib.redirect_stdout(
            output_buffer
        ):

            with contextlib.redirect_stderr(
                output_buffer
            ):

                exec(
                    code,
                    execution_namespace,
                    execution_namespace
                )

    except Exception as e:

        output_before_error = (
            output_buffer
            .getvalue()
            .strip()
        )

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

"""

        if output_before_error:

            result += f"""
📤 OUTPUT BEFORE ERROR

{output_before_error}

"""

        result += f"""
🔎 ERROR LOCATION

The error occurred around
line {line_number}.

💡 MENTOR

{explain_error(error_type)}

🛠️ POSSIBLE FIX

Review the affected line and check
your variables, data types and logic.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
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
    # PROGRAM OUTPUT
    # ========================================================

    program_output = (
        output_buffer
        .getvalue()
        .strip()
    )

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
    ) = calculate_complexity(
        code
    )

    # ========================================================
    # RECOMMENDATION
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
    # COMPLEXITY TIP
    # ========================================================

    if complexity == "LOW":

        complexity_tip = (
            "Your code is simple and easy to understand."
        )

    elif complexity == "MEDIUM":

        complexity_tip = (
            "Consider breaking large sections "
            "into smaller functions."
        )

    else:

        complexity_tip = (
            "Try reducing nesting and splitting "
            "the program into smaller functions."
        )

    # ========================================================
    # IMPROVEMENTS
    # ========================================================

    improvements = []

    try:

        tree = ast.parse(
            code
        )

        poor_names = {
            "x",
            "y",
            "z",
            "a",
            "b",
            "c"
        }

        for node in ast.walk(tree):

            if isinstance(
                node,
                ast.Name
            ):

                if node.id in poor_names:

                    improvements.append(
                        f"Consider replacing "
                        f"'{node.id}' with a "
                        "descriptive variable name."
                    )

        if (
            functions == 0
            and
            lines > 10
        ):

            improvements.append(
                "Consider using functions to divide "
                "your program into reusable components."
            )

        if (
            "input(" in code
            and
            "try:" not in code
        ):

            improvements.append(
                "Consider using try/except "
                "for input validation."
            )

        if re.findall(
            r"(?<![\w.])\d{2,}(?![\w.])",
            code
        ):

            improvements.append(
                "Consider storing frequently used "
                "numbers in named constants."
            )

    except Exception:
        pass

    # ========================================================
    # SAVE SUCCESS
    # ========================================================

    save_submission(
        current_username,
        "SUCCESS",
        code
    )

    # ========================================================
    # RESULT
    # ========================================================

    result = """
✅ CODE EXECUTED SUCCESSFULLY

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📤 PROGRAM OUTPUT

"""

    if program_output:

        result += (
            program_output
            +
            "\n\n"
        )

    else:

        result += (
            "No output produced by the program.\n\n"
        )

    result += f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🧹 CODE QUALITY

Score:
{score}/100

"""

    if issues:

        result += "Issues Found:\n\n"

        for issue in issues:

            result += (
                f"• {issue}\n"
            )

    else:

        result += (
            "🎉 No major quality issues detected.\n"
        )

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

🧑‍💻 MENTOR RECOMMENDATION

{recommendation}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

✨ AUTOMATIC IMPROVEMENTS

"""

    if improvements:

        for number, improvement in enumerate(
            improvements,
            start=1
        ):

            result += (
                f"{number}. "
                f"{improvement}\n"
            )

    else:

        result += (
            "🎉 Your code already follows "
            "several good practices!"
        )

    show_result(
        result
    )

    # ========================================================
    # VOICE
    # ========================================================

    voice_message = (
        "Code executed successfully. "
        f"Your quality score is {score} "
        "out of 100. "
        f"Complexity is {complexity}. "
        f"{recommendation}"
    )

    if program_output:

        voice_message += (
            f" The program output is "
            f"{program_output}"
        )

    speak_text(
        voice_message
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
# DASHBOARD
# ============================================================

def update_dashboard():

    if not current_username:
        return

    cursor.execute(
        """
        SELECT status
        FROM submissions
        WHERE username = ?
        """,
        (
            current_username,
        )
    )

    submissions = cursor.fetchall()

    total = len(
        submissions
    )

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

    rate = (
        successful
        /
        total
        *
        100
        if total
        else 0
    )

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
# ERROR HISTORY
# ============================================================

def show_history():

    cursor.execute(
        """
        SELECT
            error_type,
            error_message,
            line_number,
            timestamp
        FROM errors
        WHERE username = ?
        ORDER BY id DESC
        """,
        (
            current_username,
        )
    )

    records = cursor.fetchall()

    if not records:

        show_result(
            "🎉 No errors recorded yet!"
        )

        return

    text = (
        f"📜 {current_username.upper()}'S "
        "ERROR HISTORY\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    )

    for record in records:

        text += (
            f"❌ Error: {record[0]}\n"
            f"Message: {record[1]}\n"
            f"Line: {record[2]}\n"
            f"Time: {record[3]}\n\n"
            "────────────────────────────\n\n"
        )

    show_result(
        text
    )


# ============================================================
# BUG PATTERNS
# ============================================================

def show_bug_patterns():

    cursor.execute(
        """
        SELECT error_type
        FROM errors
        WHERE username = ?
        """,
        (
            current_username,
        )
    )

    records = cursor.fetchall()

    if not records:

        show_result(
            "🎉 No bug data available yet!"
        )

        return

    counts = {}

    for record in records:

        error = record[0]

        counts[error] = (
            counts.get(
                error,
                0
            )
            +
            1
        )

    sorted_errors = sorted(
        counts.items(),
        key=lambda item: item[1],
        reverse=True
    )

    most_common = (
        sorted_errors[0][0]
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
        "Practice objects and methods."
    }

    recommendation = recommendations.get(
        most_common,
        "Review your error history and keep practicing."
    )

    text = (
        f"🐛 {current_username.upper()}'S "
        "BUG PATTERNS\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    )

    for error, count in sorted_errors:

        text += (
            f"❌ {error}: "
            f"{count} time(s)\n"
        )

    text += (
        "\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "⚠️ MOST COMMON BUG\n\n"
        f"{most_common}\n\n"
        "📚 RECOMMENDATION\n\n"
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
# GRAPHS
# ============================================================

def show_graphs():

    if not MATPLOTLIB_AVAILABLE:

        messagebox.showerror(
            "Matplotlib Missing",
            "Install matplotlib using:\n\n"
            "pip install matplotlib"
        )

        return

    cursor.execute(
        """
        SELECT error_type
        FROM errors
        WHERE username = ?
        """,
        (
            current_username,
        )
    )

    records = cursor.fetchall()

    if not records:

        messagebox.showinfo(
            "No Data",
            "No error data available yet."
        )

        return

    counts = {}

    for row in records:

        error = row[0]

        counts[error] = (
            counts.get(
                error,
                0
            )
            +
            1
        )

    graph_window = tk.Toplevel(
        root
    )

    graph_window.title(
        "📊 Error Analytics"
    )

    graph_window.geometry(
        "1050x650"
    )

    graph_window.configure(
        bg="#101820"
    )

    figure = plt.Figure(
        figsize=(11, 6),
        dpi=100
    )

    # BAR

    ax1 = figure.add_subplot(
        121
    )

    names = list(
        counts.keys()
    )

    values = list(
        counts.values()
    )

    ax1.bar(
        names,
        values,
        width=0.6
    )

    ax1.set_title(
        "Error Frequency"
    )

    ax1.set_xlabel(
        "Error Type"
    )

    ax1.set_ylabel(
        "Number of Errors"
    )

    ax1.tick_params(
        axis="x",
        rotation=45
    )

    # PIE

    ax2 = figure.add_subplot(
        122
    )

    ax2.pie(
        values,
        labels=names,
        autopct="%1.1f%%",
        startangle=90
    )

    ax2.set_title(
        "Error Distribution (%)"
    )

    figure.tight_layout()

    canvas = FigureCanvasTkAgg(
        figure,
        master=graph_window
    )

    canvas.draw()

    canvas.get_tk_widget().pack(
        fill="both",
        expand=True,
        padx=15,
        pady=15
    )


# ============================================================
# PRACTICE MODE
# ============================================================

def open_practice_mode():

    global practice_difficulty
    global practice_type
    global current_question

    global practice_score
    global practice_total
    global practice_correct
    global practice_wrong
    global practice_answered

    # RESET SESSION

    practice_score = 0
    practice_total = 0
    practice_correct = 0
    practice_wrong = 0
    practice_answered = False
    current_question = None

    # WINDOW

    practice_window = tk.Toplevel(
        root
    )

    practice_window.title(
        "🎯 Practice Mode"
    )

    practice_window.geometry(
        "950x760"
    )

    practice_window.configure(
        bg="#101820"
    )

    # TITLE

    title = tk.Label(
        practice_window,
        text="🎯 PYTHON PRACTICE MODE",
        font=("Segoe UI", 22, "bold"),
        bg="#101820",
        fg="white"
    )

    title.pack(
        pady=(18, 5)
    )

    subtitle = tk.Label(
        practice_window,
        text="Choose difficulty and question type",
        font=("Segoe UI", 11),
        bg="#101820",
        fg="#aaaaaa"
    )

    subtitle.pack(
        pady=(0, 10)
    )

    # SCORE BAR

    score_frame = tk.Frame(
        practice_window,
        bg="#17232c"
    )

    score_frame.pack(
        fill="x",
        padx=25,
        pady=10
    )

    score_label = tk.Label(
        score_frame,
        text=(
            "Score: 0/0 | Correct: 0 | "
            "Wrong: 0 | Accuracy: 0%"
        ),
        font=("Segoe UI", 12, "bold"),
        bg="#17232c",
        fg="#4ade80"
    )

    score_label.pack(
        pady=10
    )

    # SELECTION

    selection_frame = tk.Frame(
        practice_window,
        bg="#101820"
    )

    selection_frame.pack(
        pady=5
    )

    tk.Label(
        selection_frame,
        text="Difficulty:",
        font=("Segoe UI", 12, "bold"),
        bg="#101820",
        fg="white"
    ).grid(
        row=0,
        column=0,
        padx=8
    )

    difficulty_var = tk.StringVar(
        value="Easy"
    )

    difficulty_box = ttk.Combobox(
        selection_frame,
        textvariable=difficulty_var,
        values=[
            "Easy",
            "Medium",
            "Tough"
        ],
        state="readonly",
        width=15
    )

    difficulty_box.grid(
        row=0,
        column=1,
        padx=8
    )

    tk.Label(
        selection_frame,
        text="Question Type:",
        font=("Segoe UI", 12, "bold"),
        bg="#101820",
        fg="white"
    ).grid(
        row=0,
        column=2,
        padx=8
    )

    type_var = tk.StringVar(
        value="MCQ"
    )

    type_box = ttk.Combobox(
        selection_frame,
        textvariable=type_var,
        values=[
            "MCQ",
            "Code Writing",
            "Code Output"
        ],
        state="readonly",
        width=18
    )

    type_box.grid(
        row=0,
        column=3,
        padx=8
    )

    # QUESTION

    question_frame = tk.Frame(
        practice_window,
        bg="#17232c"
    )

    question_frame.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=10
    )

    question_label = tk.Label(
        question_frame,
        text="Click START PRACTICE",
        font=("Consolas", 13),
        bg="#17232c",
        fg="white",
        justify="left",
        anchor="nw",
        wraplength=850
    )

    question_label.pack(
        fill="x",
        padx=20,
        pady=20
    )

    # ANSWER

    answer_frame = tk.Frame(
        question_frame,
        bg="#17232c"
    )

    answer_frame.pack(
        fill="both",
        expand=True,
        padx=20
    )

    answer_var = tk.StringVar()

    answer_entry = tk.Text(
        answer_frame,
        bg="#0b1115",
        fg="white",
        insertbackground="white",
        font=("Consolas", 12),
        height=9,
        wrap="word"
    )

    option_buttons = []

    # RESULT

    result_label_practice = tk.Label(
        question_frame,
        text="",
        font=("Segoe UI", 11, "bold"),
        bg="#17232c",
        fg="white",
        justify="left",
        anchor="w",
        wraplength=850
    )

    result_label_practice.pack(
        fill="x",
        padx=20,
        pady=10
    )

    # ========================================================
    # UPDATE SCORE
    # ========================================================

    def update_practice_score():

        accuracy = (
            practice_correct
            /
            practice_total
            *
            100
            if practice_total
            else 0
        )

        score_label.config(
            text=(
                f"Score: {practice_correct}/"
                f"{practice_total} | "
                f"Correct: {practice_correct} | "
                f"Wrong: {practice_wrong} | "
                f"Accuracy: {accuracy:.1f}%"
            )
        )

    # ========================================================
    # START PRACTICE
    # ========================================================

    def start_practice():

        global practice_difficulty
        global practice_type
        global current_question
        global practice_answered

        practice_difficulty = (
            difficulty_var.get()
        )

        practice_type = (
            type_var.get()
        )

        questions = PRACTICE_DATA[
            practice_difficulty
        ][
            practice_type
        ]

        current_question = random.choice(
            questions
        )

        practice_answered = False

        result_label_practice.config(
            text="",
            fg="white"
        )

        answer_var.set("")

        for button in option_buttons:

            button.destroy()

        option_buttons.clear()

        answer_entry.pack_forget()

        # MCQ

        if practice_type == "MCQ":

            for option in current_question[
                "options"
            ]:

                button = tk.Radiobutton(
                    answer_frame,
                    text=option,
                    variable=answer_var,
                    value=option,
                    font=("Segoe UI", 11),
                    bg="#17232c",
                    fg="white",
                    selectcolor="#24333d",
                    activebackground="#17232c",
                    activeforeground="white",
                    anchor="w"
                )

                button.pack(
                    fill="x",
                    pady=5
                )

                option_buttons.append(
                    button
                )

        # CODE

        else:

            answer_entry.delete(
                "1.0",
                tk.END
            )

            answer_entry.pack(
                fill="both",
                expand=True
            )

        question_label.config(
            text=(
                f"🎯 {practice_difficulty} | "
                f"{practice_type}\n\n"
                f"{current_question['question']}"
            )
        )

        speak_text(
            f"New {practice_difficulty} "
            f"{practice_type} question. "
            f"{current_question['question']}"
        )

    # ========================================================
    # SUBMIT
    # ========================================================

    def submit_answer():

        global practice_score
        global practice_total
        global practice_correct
        global practice_wrong
        global practice_answered

        if current_question is None:

            messagebox.showwarning(
                "No Question",
                "Click START PRACTICE first."
            )

            return

        if practice_answered:

            messagebox.showinfo(
                "Already Submitted",
                "This question is already submitted. "
                "Click NEXT QUESTION."
            )

            return

        # GET ANSWER

        if practice_type == "MCQ":

            user_answer = (
                answer_var
                .get()
                .strip()
            )

        else:

            user_answer = (
                answer_entry
                .get(
                    "1.0",
                    tk.END
                )
                .strip()
            )

        if not user_answer:

            messagebox.showwarning(
                "Answer Required",
                "Please enter/select an answer."
            )

            return

        correct_answer = (
            current_question["answer"]
        )

        # NORMALIZE

        user_normalized = (
            user_answer
            .strip()
            .lower()
            .replace(" ", "")
            .replace("\n", "")
            .replace("\r", "")
        )

        correct_normalized = (
            correct_answer
            .strip()
            .lower()
            .replace(" ", "")
            .replace("\n", "")
            .replace("\r", "")
        )

        is_correct = (
            user_normalized
            ==
            correct_normalized
        )

        # COUNT

        practice_total += 1

        practice_answered = True

        if is_correct:

            practice_correct += 1
            practice_score += 1

        else:

            practice_wrong += 1

        update_practice_score()

        accuracy = (
            practice_correct
            /
            practice_total
            *
            100
        )

        # CORRECT

        if is_correct:

            result_text = (
                "✅ CORRECT!\n\n"
                "🎯 PRACTICE ANALYSIS\n\n"
                f"Current Score: "
                f"{practice_correct}/"
                f"{practice_total}\n"
                f"Accuracy: "
                f"{accuracy:.1f}%\n\n"
                "💡 SOLUTION / EXPLANATION\n\n"
                f"{current_question['explanation']}"
            )

            result_label_practice.config(
                text=result_text,
                fg="#4ade80"
            )

            speak_text(
                "Correct answer! "
                f"Your current score is "
                f"{practice_correct} out of "
                f"{practice_total}. "
                f"Your accuracy is "
                f"{accuracy:.1f} percent. "
                "Here is the solution. "
                f"{current_question['explanation']}"
            )

        # WRONG

        else:

            result_text = (
                "❌ INCORRECT\n\n"
                "🎯 PRACTICE ANALYSIS\n\n"
                f"Current Score: "
                f"{practice_correct}/"
                f"{practice_total}\n"
                f"Accuracy: "
                f"{accuracy:.1f}%\n\n"
                "❌ YOUR ANSWER\n\n"
                f"{user_answer}\n\n"
                "✅ CORRECT SOLUTION\n\n"
                f"{correct_answer}\n\n"
                "💡 EXPLANATION\n\n"
                f"{current_question['explanation']}"
            )

            result_label_practice.config(
                text=result_text,
                fg="#ff6b6b"
            )

            speak_text(
                "That answer is incorrect. "
                f"Your current score is "
                f"{practice_correct} out of "
                f"{practice_total}. "
                f"Your accuracy is "
                f"{accuracy:.1f} percent. "
                f"The correct solution is "
                f"{correct_answer}. "
                "Here is the explanation. "
                f"{current_question['explanation']}"
            )

    # ========================================================
    # NEXT QUESTION
    # ========================================================

    def next_question():

        if current_question is None:

            start_practice()

        elif not practice_answered:

            messagebox.showwarning(
                "Submit First",
                "Please submit your current answer first."
            )

        else:

            start_practice()

    # ========================================================
    # PRACTICE ANALYSIS
    # ========================================================

    def show_practice_analysis():

        if practice_total == 0:

            messagebox.showinfo(
                "Practice Analysis",
                "Solve at least one question first."
            )

            return

        accuracy = (
            practice_correct
            /
            practice_total
            *
            100
        )

        if accuracy >= 90:

            level = "🔥 Excellent!"

            advice = (
                "You are performing very well. "
                "Try tougher questions."
            )

        elif accuracy >= 75:

            level = "⭐ Very Good!"

            advice = (
                "Good performance. "
                "Keep practicing consistently."
            )

        elif accuracy >= 50:

            level = "📚 Needs Practice"

            advice = (
                "Your basics are developing. "
                "Review incorrect questions."
            )

        else:

            level = "💪 Keep Practicing"

            advice = (
                "Focus on understanding concepts "
                "before attempting tougher questions."
            )

        analysis = (
            "🎯 PRACTICE PERFORMANCE ANALYSIS\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            f"Difficulty: {practice_difficulty}\n"
            f"Question Type: {practice_type}\n\n"
            f"Total Questions: {practice_total}\n"
            f"Correct: {practice_correct}\n"
            f"Wrong: {practice_wrong}\n"
            f"Score: {practice_correct}/"
            f"{practice_total}\n"
            f"Accuracy: {accuracy:.1f}%\n\n"
            f"Performance: {level}\n\n"
            f"💡 Mentor Advice:\n"
            f"{advice}\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )

        result_label_practice.config(
            text=analysis,
            fg="#60a5fa"
        )

        speak_text(
            "Practice analysis. "
            f"You attempted "
            f"{practice_total} questions. "
            f"You got "
            f"{practice_correct} correct "
            f"and "
            f"{practice_wrong} wrong. "
            f"Your accuracy is "
            f"{accuracy:.1f} percent. "
            f"{advice}"
        )

    # ========================================================
    # BUTTONS
    # ========================================================

    button_frame = tk.Frame(
        practice_window,
        bg="#101820"
    )

    button_frame.pack(
        fill="x",
        pady=15
    )

    start_button = tk.Button(
        button_frame,
        text="🚀 START PRACTICE",
        font=("Segoe UI", 11, "bold"),
        bg="#2d89ef",
        fg="white",
        relief="flat",
        padx=15,
        pady=9,
        command=start_practice
    )

    start_button.pack(
        side="left",
        padx=5
    )

    submit_button = tk.Button(
        button_frame,
        text="✅ SUBMIT ANSWER",
        font=("Segoe UI", 11, "bold"),
        bg="#16875a",
        fg="white",
        relief="flat",
        padx=15,
        pady=9,
        command=submit_answer
    )

    submit_button.pack(
        side="left",
        padx=5
    )

    next_button = tk.Button(
        button_frame,
        text="➡️ NEXT QUESTION",
        font=("Segoe UI", 11, "bold"),
        bg="#7c3aed",
        fg="white",
        relief="flat",
        padx=15,
        pady=9,
        command=next_question
    )

    next_button.pack(
        side="left",
        padx=5
    )

    analysis_button = tk.Button(
        button_frame,
        text="📊 PRACTICE ANALYSIS",
        font=("Segoe UI", 11, "bold"),
        bg="#d97706",
        fg="white",
        relief="flat",
        padx=15,
        pady=9,
        command=show_practice_analysis
    )

    analysis_button.pack(
        side="left",
        padx=5
    )


# ============================================================
# LOGIN
# ============================================================

def login():

    global current_username

    username = (
        username_entry
        .get()
        .strip()
    )

    if not username:

        messagebox.showwarning(
            "Username Required",
            "Please enter your username."
        )

        return

    cursor.execute(
        """
        SELECT username
        FROM users
        WHERE username = ?
        """,
        (
            username,
        )
    )

    user = cursor.fetchone()

    if user is None:

        cursor.execute(
            """
            INSERT INTO users (username)
            VALUES (?)
            """,
            (
                username,
            )
        )

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
    "1250x800"
)

root.minsize(
    1050,
    700
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

login_subtitle.pack()


username_label = tk.Label(
    login_frame,
    text="👤 Username",
    font=("Segoe UI", 12, "bold"),
    bg="#101820",
    fg="white"
)

username_label.pack(
    pady=(25, 5)
)


username_entry = tk.Entry(
    login_frame,
    font=("Segoe UI", 14),
    width=30,
    justify="center"
)

username_entry.pack(
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
    pady=20
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
    width=205
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


def sidebar_button(
        text,
        command,
        active=False):

    button = tk.Button(
        sidebar,
        text=text,
        font=("Segoe UI", 10),
        bg=(
            "#2d89ef"
            if active
            else
            "#24333d"
        ),
        fg="white",
        relief="flat",
        anchor="w",
        padx=12,
        pady=8,
        command=command
    )

    button.pack(
        fill="x",
        padx=12,
        pady=4
    )

    return button


analyze_button = sidebar_button(
    "💻 Analyze Code",
    lambda: code_editor.focus(),
    True
)

history_button = sidebar_button(
    "📜 Error History",
    show_history
)

bugs_button = sidebar_button(
    "🐛 Bug Patterns",
    show_bug_patterns
)

graph_button = sidebar_button(
    "📊 Error Graphs",
    show_graphs
)

practice_button = sidebar_button(
    "🎯 Practice Mode",
    open_practice_mode
)

voice_button = sidebar_button(
    "🔊 Voice ON",
    toggle_voice
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


def create_card(
        parent,
        title):

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

    title_label_card = tk.Label(
        frame,
        text=title,
        font=("Segoe UI", 9),
        bg="#17232c",
        fg="#aaaaaa"
    )

    title_label_card.pack(
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
# CODE EDITOR
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


editor_frame = tk.Frame(
    right_area,
    bg="#17232c",
    height=280
)

editor_frame.pack(
    fill="x",
    pady=(5, 10)
)

editor_frame.pack_propagate(
    False
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
# BUTTONS
# ============================================================

button_frame = tk.Frame(
    right_area,
    bg="#101820"
)

button_frame.pack(
    fill="x",
    pady=(0, 10)
)


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
# RESULT
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


result_frame = tk.Frame(
    right_area,
    bg="#17232c"
)

result_frame.pack(
    fill="both",
    expand=True
)


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
# WINDOW CLOSE
# ============================================================

def close_application():

    try:

        connection.commit()
        connection.close()

    except Exception:
        pass

    root.destroy()


root.protocol(
    "WM_DELETE_WINDOW",
    close_application
)


# ============================================================
# START
# ============================================================

root.mainloop()