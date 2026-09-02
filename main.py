import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import traceback
import ast
import re
import random

# ============================================================
# OPTIONAL WINDOWS VOICE
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
# OPTIONAL MATPLOTLIB
# ============================================================

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False


# ============================================================
# DATABASE
# ============================================================

connection = sqlite3.connect("coding_history.db")
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
practice_attempts = 0
practice_correct = 0
practice_wrong = 0


# ============================================================
# PRACTICE QUESTION DATABASE
# ============================================================

PRACTICE_DATA = {

    "Easy": {

        "MCQ": [
            {
                "question": "What is the output of: print(2 + 3 * 2)?",
                "options": ["10", "8", "12", "7"],
                "answer": "8",
                "explanation":
                    "Multiplication happens before addition. "
                    "3 * 2 = 6 and 2 + 6 = 8.",
                "solution":
                    "First calculate 3 * 2 = 6. Then 2 + 6 = 8. "
                    "Therefore, the output is 8."
            },
            {
                "question": "Which keyword is used to define a function?",
                "options": ["function", "define", "def", "fun"],
                "answer": "def",
                "explanation":
                    "Python uses the def keyword to create a function.",
                "solution":
                    "The correct keyword is def. Example: def add(a, b):"
            },
            {
                "question": "Which data type stores True or False?",
                "options": ["int", "str", "bool", "float"],
                "answer": "bool",
                "explanation":
                    "Boolean values in Python are True and False.",
                "solution":
                    "The bool data type represents True or False values."
            },
            {
                "question": "What does len([10, 20, 30]) return?",
                "options": ["2", "3", "4", "30"],
                "answer": "3",
                "explanation":
                    "The list contains three elements.",
                "solution":
                    "The list has 3 elements: 10, 20 and 30. Therefore len() returns 3."
            }
        ],

        "Code Writing": [
            {
                "question": "Write Python code to print 'Hello World'.",
                "answer": "print('Hello World')",
                "explanation": "Use Python's print() function.",
                "solution":
                    "Use the print() function:\n\nprint('Hello World')"
            },
            {
                "question":
                    "Write Python code to calculate the sum of two numbers "
                    "stored in variables a and b.",
                "answer": "print(a + b)",
                "explanation":
                    "The + operator adds the values of a and b.",
                "solution":
                    "Use the + operator:\n\nprint(a + b)"
            }
        ],

        "Code Output": [
            {
                "question":
                    "What is the output?\n\n"
                    "x = 5\n"
                    "y = 2\n"
                    "print(x + y)",
                "answer": "7",
                "explanation": "5 + 2 = 7.",
                "solution":
                    "x = 5 and y = 2.\n"
                    "The expression x + y becomes 5 + 2 = 7.\n"
                    "Output: 7"
            },
            {
                "question":
                    "What is the output?\n\n"
                    "name = 'Python'\n"
                    "print(len(name))",
                "answer": "6",
                "explanation": "'Python' contains 6 characters.",
                "solution":
                    "'Python' has six characters: P, y, t, h, o, n.\n"
                    "Output: 6"
            }
        ]
    },

    "Medium": {

        "MCQ": [
            {
                "question":
                    "What is the output?\n"
                    "numbers = [1, 2, 3]\n"
                    "print(numbers[-1])",
                "options": ["1", "2", "3", "Error"],
                "answer": "3",
                "explanation":
                    "Index -1 refers to the last element of a list.",
                "solution":
                    "Negative index -1 means the last item of the list.\n"
                    "Therefore numbers[-1] is 3."
            },
            {
                "question":
                    "Which method adds an element to the end of a list?",
                "options": ["add()", "append()", "insert_end()", "push()"],
                "answer": "append()",
                "explanation":
                    "list.append(value) adds a value to the end.",
                "solution":
                    "The append() method adds an item at the end of a list."
            },
            {
                "question":
                    "What does range(5) generate?",
                "options":
                    ["1,2,3,4,5", "0,1,2,3,4", "0,1,2,3,4,5", "5 only"],
                "answer": "0,1,2,3,4",
                "explanation":
                    "range(5) starts at 0 and stops before 5.",
                "solution":
                    "range(5) produces values starting at 0 and ending before 5:\n"
                    "0, 1, 2, 3, 4."
            }
        ],

        "Code Writing": [
            {
                "question":
                    "Write a Python program that prints numbers "
                    "from 1 to 5 using a for loop.",
                "answer":
                    "for i in range(1, 6):\n    print(i)",
                "explanation":
                    "range(1, 6) produces 1, 2, 3, 4, 5.",
                "solution":
                    "Use range(1, 6) because the ending value 6 is excluded:\n\n"
                    "for i in range(1, 6):\n"
                    "    print(i)"
            },
            {
                "question":
                    "Write a Python program to check whether a number "
                    "is even or odd.",
                "answer":
                    "if n % 2 == 0:\n    print('Even')\nelse:\n    print('Odd')",
                "explanation":
                    "A number is even when its remainder after division "
                    "by 2 is zero.",
                "solution":
                    "Use the modulus operator %:\n\n"
                    "if n % 2 == 0:\n"
                    "    print('Even')\n"
                    "else:\n"
                    "    print('Odd')"
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
                "answer": "6",
                "explanation":
                    "The loop calculates 1 + 2 + 3 = 6.",
                "solution":
                    "The loop runs for i = 1, 2 and 3.\n"
                    "total becomes 1, then 3, then 6.\n"
                    "Output: 6"
            },
            {
                "question":
                    "What is the output?\n\n"
                    "numbers = [10, 20, 30]\n"
                    "numbers.append(40)\n"
                    "print(len(numbers))",
                "answer": "4",
                "explanation":
                    "append() adds 40, so the list contains four elements.",
                "solution":
                    "The original list has 3 elements. append(40) adds one more.\n"
                    "Therefore len(numbers) is 4."
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
                "options":
                    ["[1, 2, 3]", "[1, 2, 3, 4]", "[4]", "Error"],
                "answer": "[1, 2, 3, 4]",
                "explanation":
                    "y references the same list object as x.",
                "solution":
                    "y = x does not create a new list. Both variables refer "
                    "to the same list. Therefore append(4) changes x too."
            },
            {
                "question":
                    "Which concept allows a function to call itself?",
                "options":
                    ["Iteration", "Inheritance", "Recursion", "Encapsulation"],
                "answer": "Recursion",
                "explanation":
                    "Recursion occurs when a function calls itself.",
                "solution":
                    "A function calling itself is called recursion."
            },
            {
                "question":
                    "What is the average-case lookup complexity of a Python dictionary?",
                "options":
                    ["O(n)", "O(log n)", "O(1)", "O(n²)"],
                "answer": "O(1)",
                "explanation":
                    "Python dictionaries use hash tables, giving average O(1) lookup.",
                "solution":
                    "Python dictionaries are hash-table based. Their average lookup "
                    "time is O(1)."
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
                    "whenever a bigger value is found.",
                "solution":
                    "A simple solution is:\n\n"
                    "largest = numbers[0]\n"
                    "for number in numbers:\n"
                    "    if number > largest:\n"
                    "        largest = number\n"
                    "print(largest)"
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
                    "The dictionary stores each item and increments its count.",
                "solution":
                    "Use a dictionary and update the count for every item:\n\n"
                    "frequency = {}\n"
                    "for item in numbers:\n"
                    "    frequency[item] = frequency.get(item, 0) + 1"
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
                "answer": "24",
                "explanation":
                    "The recursive function calculates 4 × 3 × 2 × 1 = 24.",
                "solution":
                    "fun(4) = 4 * fun(3)\n"
                    "fun(3) = 3 * fun(2)\n"
                    "fun(2) = 2 * fun(1)\n"
                    "fun(1) = 1 * fun(0)\n"
                    "fun(0) = 1\n\n"
                    "Therefore 4 × 3 × 2 × 1 = 24.\n"
                    "Output: 24"
            },
            {
                "question":
                    "What is the output?\n\n"
                    "data = {'a': 1, 'b': 2}\n"
                    "data['c'] = data['a'] + data['b']\n"
                    "print(data['c'])",
                "answer": "3",
                "explanation":
                    "data['a'] + data['b'] = 1 + 2 = 3.",
                "solution":
                    "data['a'] is 1 and data['b'] is 2.\n"
                    "So data['c'] becomes 1 + 2 = 3.\n"
                    "Output: 3"
            }
        ]
    }
}


# ============================================================
# DATABASE FUNCTIONS
# ============================================================

def save_submission(username, status, code):
    cursor.execute("""
    INSERT INTO submissions (username, status, code)
    VALUES (?, ?, ?)
    """, (username, status, code))
    connection.commit()


def save_error(username, error_type, error_message, line_number, code):
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
# QUALITY ANALYZER
# ============================================================

def calculate_quality(code):
    lines = code.splitlines()
    score = 100
    issues = []

    for number, line in enumerate(lines, start=1):
        if len(line) > 80:
            issues.append(
                f"Line {number}: Line longer than 80 characters."
            )
            score -= 5

    poor_names = {"x", "y", "z", "a", "b", "c", "q"}

    for number, line in enumerate(lines, start=1):
        stripped = line.strip()

        if "=" in stripped and not stripped.startswith("#"):
            variable = stripped.split("=")[0].strip()

            if variable in poor_names:
                issues.append(
                    f"Line {number}: Variable '{variable}' "
                    "has a vague name."
                )
                score -= 5

    for number, line in enumerate(lines, start=1):
        if "TODO" in line.upper() or "FIXME" in line.upper():
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
            f"Code contains {print_count} print statements."
        )
        score -= 5

    for number, line in enumerate(lines, start=1):
        if not line.strip():
            continue

        spaces = len(line) - len(line.lstrip())

        if spaces // 4 >= 4:
            issues.append(
                f"Line {number}: Deep nesting detected."
            )
            score -= 5
            break

    score = max(0, score)
    return score, issues


# ============================================================
# COMPLEXITY ANALYZER
# ============================================================

def calculate_complexity(code):
    lines = code.splitlines()

    actual_lines = [
        line for line in lines
        if line.strip()
    ]

    loops = 0
    conditions = 0
    functions = 0
    max_nesting = 0

    for line in actual_lines:
        stripped = line.strip()

        if stripped.startswith("for ") or stripped.startswith("while "):
            loops += 1

        if (
            stripped.startswith("if ")
            or stripped.startswith("elif ")
            or stripped.startswith("else:")
        ):
            conditions += 1

        if stripped.startswith("def "):
            functions += 1

    for line in lines:
        if not line.strip():
            continue

        spaces = len(line) - len(line.lstrip())

        max_nesting = max(
            max_nesting,
            spaces // 4
        )

    points = (
        loops * 2
        + conditions * 2
        + functions
        + max_nesting * 2
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
# RESULT HELPERS
# ============================================================

def clear_result():
    result_text.config(state="normal")
    result_text.delete("1.0", tk.END)
    result_text.config(state="disabled")


def show_result(text):
    result_text.config(state="normal")
    result_text.delete("1.0", tk.END)
    result_text.insert(tk.END, text)
    result_text.config(state="disabled")


# ============================================================
# RUN & ANALYZE
# ============================================================

def analyze_code():
    code = code_editor.get("1.0", tk.END).strip()

    if not code:
        messagebox.showwarning(
            "No Code",
            "Please enter Python code first."
        )
        return

    clear_result()

    # ---------------- SYNTAX ----------------

    try:
        compile(code, "<student_code>", "exec")

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

Check your brackets, colons,
quotes and indentation.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

        show_result(result)

        speak_text(
            f"Syntax error detected on line {e.lineno}. "
            f"{explain_error('SyntaxError')}"
        )

        update_dashboard()
        return

    # ---------------- EXECUTION ----------------

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
            tb = traceback.extract_tb(e.__traceback__)

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

🔎 ERROR LOCATION

The error occurred around line {line_number}.

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

    # ---------------- QUALITY ----------------

    score, issues = calculate_quality(code)

    # ---------------- COMPLEXITY ----------------

    (
        complexity,
        points,
        lines,
        loops,
        conditions,
        functions,
        nesting
    ) = calculate_complexity(code)

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

    improvements = []

    try:
        tree = ast.parse(code)

        poor_names = {
            "x", "y", "z", "a", "b", "c"
        }

        found_names = set()

        for node in ast.walk(tree):
            if isinstance(node, ast.Name):
                if node.id in poor_names:
                    found_names.add(node.id)

        for name in sorted(found_names):
            improvements.append(
                f"Consider replacing '{name}' "
                "with a descriptive variable name."
            )

        if functions == 0 and lines > 10:
            improvements.append(
                "Consider using functions to divide "
                "your program into reusable components."
            )

        if "input(" in code and "try:" not in code:
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

    save_submission(
        current_username,
        "SUCCESS",
        code
    )

    result = f"""
✅ CODE EXECUTED SUCCESSFULLY

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🧹 CODE QUALITY

Score:
{score}/100

"""

    if issues:
        result += "Issues Found:\n\n"

        for issue in issues:
            result += f"• {issue}\n"
    else:
        result += "🎉 No major quality issues detected.\n"

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
            result += f"{number}. {improvement}\n"
    else:
        result += (
            "🎉 Your code already follows "
            "several good practices!"
        )

    show_result(result)

    speak_text(
        f"Code executed successfully. "
        f"Your quality score is {score} out of 100. "
        f"Complexity is {complexity}. "
        f"Mentor recommendation: {recommendation}"
    )

    update_dashboard()


# ============================================================
# CLEAR CODE
# ============================================================

def clear_code():
    code_editor.delete("1.0", tk.END)
    clear_result()


# ============================================================
# DASHBOARD
# ============================================================

def update_dashboard():

    if not current_username:
        return

    cursor.execute("""
    SELECT status
    FROM submissions
    WHERE username = ?
    """, (current_username,))

    submissions = cursor.fetchall()

    total = len(submissions)

    successful = sum(
        1 for row in submissions
        if row[0] == "SUCCESS"
    )

    failed = sum(
        1 for row in submissions
        if row[0] == "FAILED"
    )

    rate = (
        successful / total * 100
        if total
        else 0
    )

    total_label.config(text=str(total))
    success_label.config(text=str(successful))
    failed_label.config(text=str(failed))
    rate_label.config(text=f"{rate:.1f}%")


# ============================================================
# ERROR HISTORY
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
    """, (current_username,))

    records = cursor.fetchall()

    if not records:
        show_result("🎉 No errors recorded yet!")
        return

    text = (
        f"📜 {current_username.upper()}'S ERROR HISTORY\n\n"
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

    show_result(text)


# ============================================================
# BUG PATTERNS
# ============================================================

def show_bug_patterns():

    cursor.execute("""
    SELECT error_type
    FROM errors
    WHERE username = ?
    """, (current_username,))

    records = cursor.fetchall()

    if not records:
        show_result("🎉 No bug data available yet!")
        return

    counts = {}

    for record in records:
        error = record[0]
        counts[error] = counts.get(error, 0) + 1

    sorted_errors = sorted(
        counts.items(),
        key=lambda item: item[1],
        reverse=True
    )

    most_common = sorted_errors[0][0]

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
        f"🐛 {current_username.upper()}'S BUG PATTERNS\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    )

    for error, count in sorted_errors:
        text += f"❌ {error}: {count} time(s)\n"

    text += (
        "\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "⚠️ MOST COMMON BUG\n\n"
        f"{most_common}\n\n"
        "📚 RECOMMENDATION\n\n"
        f"{recommendation}"
    )

    show_result(text)

    speak_text(
        f"Your most common error is "
        f"{most_common}. "
        f"Recommendation: {recommendation}"
    )


# ============================================================
# ERROR GRAPHS
# ============================================================

def show_graphs():

    if not MATPLOTLIB_AVAILABLE:
        messagebox.showerror(
            "Matplotlib Missing",
            "Install matplotlib using:\n\npip install matplotlib"
        )
        return

    cursor.execute("""
    SELECT error_type
    FROM errors
    WHERE username = ?
    """, (current_username,))

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
        counts[error] = counts.get(error, 0) + 1

    graph_window = tk.Toplevel(root)
    graph_window.title("📊 Error Analytics")
    graph_window.geometry("1050x650")
    graph_window.configure(bg="#101820")

    figure = plt.Figure(
        figsize=(11, 6),
        dpi=100
    )

    ax1 = figure.add_subplot(121)

    names = list(counts.keys())
    values = list(counts.values())

    ax1.bar(names, values, width=0.6)
    ax1.set_title("Error Frequency")
    ax1.set_xlabel("Error Type")
    ax1.set_ylabel("Number of Errors")
    ax1.tick_params(axis="x", rotation=45)

    ax2 = figure.add_subplot(122)

    ax2.pie(
        values,
        labels=names,
        autopct="%1.1f%%",
        startangle=90
    )

    ax2.set_title("Error Distribution (%)")

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
# FIXED:
# 1. NEXT QUESTION BUTTON ALWAYS VISIBLE
# 2. SOLUTION SHOWN AFTER SUBMIT
# 3. SOLUTION SPOKEN BY VOICE
# 4. SCORE / ACCURACY TRACKING
# 5. NEXT QUESTION WORKS AFTER WRONG ANSWER
# ============================================================

def open_practice_mode():

    global practice_score
    global practice_attempts
    global practice_correct
    global practice_wrong

    # Start a fresh practice session
    practice_score = 0
    practice_attempts = 0
    practice_correct = 0
    practice_wrong = 0

    practice_window = tk.Toplevel(root)
    practice_window.title("🎯 Practice Mode")
    practice_window.geometry("1050x760")
    practice_window.minsize(900, 650)
    practice_window.configure(bg="#101820")

    # --------------------------------------------------------
    # HEADER
    # --------------------------------------------------------

    title = tk.Label(
        practice_window,
        text="🎯 PYTHON PRACTICE MODE",
        font=("Segoe UI", 24, "bold"),
        bg="#101820",
        fg="white"
    )

    title.pack(pady=(18, 5))

    subtitle = tk.Label(
        practice_window,
        text="Choose difficulty and question type",
        font=("Segoe UI", 11),
        bg="#101820",
        fg="#aaaaaa"
    )

    subtitle.pack(pady=(0, 10))

    # --------------------------------------------------------
    # SCORE BAR
    # --------------------------------------------------------

    score_frame = tk.Frame(
        practice_window,
        bg="#17232c",
        height=48
    )

    score_frame.pack(
        fill="x",
        padx=15,
        pady=(5, 12)
    )

    score_frame.pack_propagate(False)

    score_label = tk.Label(
        score_frame,
        text="Score: 0/0 | Correct: 0 | Wrong: 0 | Accuracy: 0.0%",
        font=("Segoe UI", 11, "bold"),
        bg="#17232c",
        fg="#4ade80"
    )

    score_label.pack(expand=True)

    # --------------------------------------------------------
    # SELECTIONS
    # --------------------------------------------------------

    selection_frame = tk.Frame(
        practice_window,
        bg="#101820"
    )

    selection_frame.pack(pady=2)

    tk.Label(
        selection_frame,
        text="Difficulty:",
        font=("Segoe UI", 11, "bold"),
        bg="#101820",
        fg="white"
    ).pack(side="left", padx=(0, 8))

    difficulty_var = tk.StringVar(value="Easy")

    difficulty_box = ttk.Combobox(
        selection_frame,
        textvariable=difficulty_var,
        values=["Easy", "Medium", "Tough"],
        state="readonly",
        width=15,
        font=("Segoe UI", 10)
    )

    difficulty_box.pack(side="left", padx=(0, 25))

    tk.Label(
        selection_frame,
        text="Question Type:",
        font=("Segoe UI", 11, "bold"),
        bg="#101820",
        fg="white"
    ).pack(side="left", padx=(0, 8))

    type_var = tk.StringVar(value="MCQ")

    type_box = ttk.Combobox(
        selection_frame,
        textvariable=type_var,
        values=["MCQ", "Code Writing", "Code Output"],
        state="readonly",
        width=18,
        font=("Segoe UI", 10)
    )

    type_box.pack(side="left")

    # --------------------------------------------------------
    # QUESTION AREA
    # IMPORTANT: FIXED HEIGHT SO BUTTONS DON'T GO BELOW WINDOW
    # --------------------------------------------------------

    question_frame = tk.Frame(
        practice_window,
        bg="#17232c",
        height=475
    )

    question_frame.pack(
        fill="x",
        padx=15,
        pady=12
    )

    question_frame.pack_propagate(False)

    question_label = tk.Label(
        question_frame,
        text="Select options and click START PRACTICE",
        font=("Consolas", 12),
        bg="#17232c",
        fg="white",
        justify="left",
        anchor="nw",
        wraplength=960
    )

    question_label.pack(
        fill="x",
        padx=20,
        pady=(18, 10)
    )

    answer_frame = tk.Frame(
        question_frame,
        bg="#17232c"
    )

    answer_frame.pack(
        fill="both",
        expand=True,
        padx=20,
        pady=(0, 5)
    )

    answer_var = tk.StringVar()

    answer_entry = tk.Text(
        answer_frame,
        bg="#0b1115",
        fg="white",
        insertbackground="white",
        font=("Consolas", 12),
        height=8,
        wrap="word",
        relief="solid",
        bd=1
    )

    option_buttons = []

    # --------------------------------------------------------
    # RESULT / SOLUTION AREA
    # --------------------------------------------------------

    result_label_practice = tk.Label(
        question_frame,
        text="",
        font=("Segoe UI", 10, "bold"),
        bg="#17232c",
        fg="white",
        justify="left",
        anchor="nw",
        wraplength=960
    )

    result_label_practice.pack(
        fill="x",
        padx=20,
        pady=(5, 10)
    )

    # --------------------------------------------------------
    # UPDATE SCORE
    # --------------------------------------------------------

    def update_practice_score():

        total = practice_attempts

        accuracy = (
            practice_correct / total * 100
            if total
            else 0
        )

        score_label.config(
            text=(
                f"Score: {practice_correct}/{total} | "
                f"Correct: {practice_correct} | "
                f"Wrong: {practice_wrong} | "
                f"Accuracy: {accuracy:.1f}%"
            )
        )

    # --------------------------------------------------------
    # START QUESTION
    # --------------------------------------------------------

    def start_practice():

        global practice_difficulty
        global practice_type
        global current_question

        practice_difficulty = difficulty_var.get()
        practice_type = type_var.get()

        questions = PRACTICE_DATA[
            practice_difficulty
        ][
            practice_type
        ]

        current_question = random.choice(questions)

        result_label_practice.config(
            text="",
            fg="white"
        )

        answer_var.set("")

        for button in option_buttons:
            button.destroy()

        option_buttons.clear()

        answer_entry.delete(
            "1.0",
            tk.END
        )

        if practice_type == "MCQ":

            answer_entry.pack_forget()

            for option in current_question["options"]:

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
                    pady=4
                )

                option_buttons.append(button)

        else:

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
            f"{practice_type} question."
        )

    # --------------------------------------------------------
    # CHECK CODE WRITING ANSWER
    # More flexible than exact string matching
    # --------------------------------------------------------

    def code_answers_match(user_answer, correct_answer):

        user_normalized = (
            user_answer
            .strip()
            .lower()
            .replace(" ", "")
            .replace("\r", "")
            .replace("\n", "")
            .replace('"', "'")
        )

        correct_normalized = (
            correct_answer
            .strip()
            .lower()
            .replace(" ", "")
            .replace("\r", "")
            .replace("\n", "")
            .replace('"', "'")
        )

        if user_normalized == correct_normalized:
            return True

        try:
            user_tree = ast.parse(user_answer)
            correct_tree = ast.parse(correct_answer)

            return (
                ast.dump(user_tree, include_attributes=False)
                == ast.dump(correct_tree, include_attributes=False)
            )

        except Exception:
            return False

    # --------------------------------------------------------
    # SUBMIT ANSWER
    # --------------------------------------------------------

    def submit_answer():

        global current_question
        global practice_attempts
        global practice_correct
        global practice_wrong

        if current_question is None:

            messagebox.showwarning(
                "No Question",
                "Click START PRACTICE first."
            )

            return

        if practice_type == "MCQ":

            user_answer = answer_var.get().strip()

        else:

            user_answer = answer_entry.get(
                "1.0",
                tk.END
            ).strip()

        if not user_answer:

            messagebox.showwarning(
                "Answer Required",
                "Please enter/select an answer."
            )

            return

        correct_answer = current_question["answer"]

        if practice_type == "Code Writing":
            is_correct = code_answers_match(
                user_answer,
                correct_answer
            )
        else:
            user_normalized = (
                user_answer
                .strip()
                .lower()
                .replace(" ", "")
                .replace("\r", "")
                .replace("\n", "")
            )

            correct_normalized = (
                correct_answer
                .strip()
                .lower()
                .replace(" ", "")
                .replace("\r", "")
                .replace("\n", "")
            )

            is_correct = (
                user_normalized
                == correct_normalized
            )

        # Count attempt
        practice_attempts += 1

        if is_correct:

            practice_correct += 1

            result_label_practice.config(
                text=(
                    "✅ CORRECT!\n\n"
                    "💡 EXPLANATION\n"
                    f"{current_question['explanation']}\n\n"
                    "📝 SOLUTION\n"
                    f"{current_question['solution']}"
                ),
                fg="#4ade80"
            )

            speak_text(
                "Correct answer. "
                + current_question["explanation"]
                + " The solution is. "
                + current_question["solution"]
            )

        else:

            practice_wrong += 1

            result_label_practice.config(
                text=(
                    "❌ INCORRECT\n\n"
                    "❌ YOUR ANSWER\n"
                    f"{user_answer}\n\n"
                    "☑ CORRECT ANSWER\n"
                    f"{correct_answer}\n\n"
                    "💡 EXPLANATION\n"
                    f"{current_question['explanation']}\n\n"
                    "📝 SOLUTION\n"
                    f"{current_question['solution']}"
                ),
                fg="#ff6b6b"
            )

            speak_text(
                "That answer is incorrect. "
                f"The correct answer is {correct_answer}. "
                + current_question["explanation"]
                + " The solution is. "
                + current_question["solution"]
            )

        update_practice_score()

    # --------------------------------------------------------
    # NEXT QUESTION
    # --------------------------------------------------------

    def next_question():

        # This is the important fix:
        # next button directly generates another question.
        start_practice()

    # --------------------------------------------------------
    # BUTTON AREA
    # This is OUTSIDE the expanding question frame, so it
    # remains visible after an incorrect answer.
    # --------------------------------------------------------

    practice_button_frame = tk.Frame(
        practice_window,
        bg="#101820",
        height=58
    )

    practice_button_frame.pack(
        fill="x",
        padx=15,
        pady=(0, 10)
    )

    practice_button_frame.pack_propagate(False)

    start_button = tk.Button(
        practice_button_frame,
        text="🚀 START PRACTICE",
        font=("Segoe UI", 10, "bold"),
        bg="#2d89ef",
        fg="white",
        relief="flat",
        padx=18,
        pady=8,
        command=start_practice
    )

    start_button.pack(
        side="left",
        padx=5
    )

    submit_button = tk.Button(
        practice_button_frame,
        text="✅ SUBMIT ANSWER",
        font=("Segoe UI", 10, "bold"),
        bg="#16875a",
        fg="white",
        relief="flat",
        padx=18,
        pady=8,
        command=submit_answer
    )

    submit_button.pack(
        side="left",
        padx=5
    )

    next_button = tk.Button(
        practice_button_frame,
        text="➡️ NEXT QUESTION",
        font=("Segoe UI", 10, "bold"),
        bg="#7c3aed",
        fg="white",
        relief="flat",
        padx=18,
        pady=8,
        command=next_question
    )

    next_button.pack(
        side="left",
        padx=5
    )

    # --------------------------------------------------------
    # CLOSE BUTTON
    # --------------------------------------------------------

    close_button = tk.Button(
        practice_button_frame,
        text="✖ CLOSE",
        font=("Segoe UI", 10),
        bg="#24333d",
        fg="white",
        relief="flat",
        padx=18,
        pady=8,
        command=practice_window.destroy
    )

    close_button.pack(
        side="right",
        padx=5
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
    """, (username,))

    user = cursor.fetchone()

    if user is None:

        cursor.execute("""
        INSERT INTO users (username)
        VALUES (?)
        """, (username,))

        connection.commit()

        welcome = (
            f"Welcome to AI Code Mentor, {username}!"
        )

    else:

        welcome = (
            f"Welcome back, {username}!"
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

    speak_text(welcome)


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
    style.theme_use("clam")
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

header.pack_propagate(False)

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

sidebar.pack_propagate(False)

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
    active=False
):

    button = tk.Button(
        sidebar,
        text=text,
        font=("Segoe UI", 10),
        bg="#2d89ef" if active else "#24333d",
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

    frame.pack_propagate(False)

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

editor_frame.pack_propagate(False)

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
