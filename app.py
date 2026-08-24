from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename
from pathlib import Path
from datetime import datetime
import sqlite3
import os

BASE_DIR = Path(__file__).resolve().parent

DB_PATH = BASE_DIR / "instance" / "vijay_education.db"
UPLOAD_DIR = BASE_DIR / "uploads"

ALLOWED_EXTENSIONS = {
    "pdf",
    "doc",
    "docx"
}

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "change-this-secret-key"
)

app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024
app.config["UPLOAD_FOLDER"] = str(UPLOAD_DIR)

UPLOAD_DIR.mkdir(
    parents=True,
    exist_ok=True
)

DB_PATH.parent.mkdir(
    parents=True,
    exist_ok=True
)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()

    conn.executescript("""
    CREATE TABLE IF NOT EXISTS enquiries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT NOT NULL,
        interested_in TEXT NOT NULL,
        message TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS faculty_applications (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT NOT NULL,
        role TEXT NOT NULL,
        cv_filename TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS contact_messages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT,
        message TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    """)

    conn.commit()
    conn.close()


def allowed_file(filename):
    return (
        "." in filename
        and
        filename.rsplit(".", 1)[1].lower()
        in ALLOWED_EXTENSIONS
    )


COURSES = [
    {
        "id": 1,
        "category": "FOCUS",
        "title": "Entrance Exam Preparation",
        "description": "A structured, mentor-led pathway for students preparing to enter their next academic chapter with confidence.",
        "duration": "6–12 months",
        "start": "January 2026",
        "fee": "₹24,000",
        "classes": "Classes 10–12",
        "includes": [
            "Weekly live classes",
            "Mock tests & progress reviews",
            "Personal academic guidance"
        ]
    },
    {
        "id": 2,
        "category": "BUILD",
        "title": "Foundation Courses",
        "description": "Strong concept-first learning designed to build academic fundamentals and long-term confidence.",
        "duration": "3–12 months",
        "start": "Flexible batches",
        "fee": "₹18,000",
        "classes": "Classes 6–10",
        "includes": [
            "Concept-focused lessons",
            "Practice assignments",
            "Mentor support"
        ]
    }
]


SITE_INFO = {
    "name": "VIJAY",
    "full_name": "VIJAY EDUCATIONAL SERVICES",
    "phone": "+91 98765 43210",
    "email": "info@vijayeducationalservices.com",
    "address": "Vijay Educational Services, Vijayawada, Andhra Pradesh, India",
    "working_hours": "Monday – Saturday | 9:00 AM – 7:00 PM",
    "facebook": "#",
    "instagram": "#",
    "youtube": "#",
    "whatsapp": "https://wa.me/919876543210",
    "map_url": "https://www.google.com/maps?q=Vijayawada,Andhra+Pradesh&output=embed"
}


@app.context_processor
def inject_globals():
    return {
        "courses": COURSES,
        "site": SITE_INFO,
        "year": datetime.now().year
    }


@app.route("/")
def home():
    return render_template("loading.html")


@app.route("/intro")
def intro():
    return render_template("intro-video.html")


@app.route("/home")
def main_home():
    return render_template("index.html")


@app.route("/about")
def about():
    return render_template("about.html")


@app.route("/courses")
def courses():
    return render_template("courses.html")


@app.route("/careers")
def careers():
    return render_template("careers.html")


@app.route("/contact")
def contact():
    return render_template("contact.html")


@app.post("/enroll")
def enroll():

    data = request.form

    required = [
        "full_name",
        "email",
        "phone",
        "interested_in"
    ]

    if not all(
        data.get(field, "").strip()
        for field in required
    ):

        flash(
            "Please complete all required fields.",
            "error"
        )

        return redirect(
            request.referrer
            or
            url_for("main_home")
        )

    conn = get_db()

    conn.execute(
        """
        INSERT INTO enquiries
        (
            full_name,
            email,
            phone,
            interested_in,
            message,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        (
            data["full_name"].strip(),
            data["email"].strip(),
            data["phone"].strip(),
            data["interested_in"].strip(),
            data.get("message", "").strip(),
            datetime.now().isoformat(
                timespec="seconds"
            )
        )
    )

    conn.commit()
    conn.close()

    flash(
        "Thanks! Your interest has been registered. Our team will contact you soon.",
        "success"
    )

    return redirect(
        request.referrer
        or
        url_for("main_home")
    )


@app.post("/faculty/apply")
def faculty_apply():

    data = request.form

    required = [
        "full_name",
        "email",
        "phone",
        "role"
    ]

    if not all(
        data.get(field, "").strip()
        for field in required
    ):

        flash(
            "Please complete all required faculty fields.",
            "error"
        )

        return redirect(
            url_for("careers")
        )

    file = request.files.get("cv")

    saved_name = None

    if file and file.filename:

        if not allowed_file(file.filename):

            flash(
                "CV must be PDF, DOC, or DOCX.",
                "error"
            )

            return redirect(
                url_for("careers")
            )

        original = secure_filename(
            file.filename
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d%H%M%S"
        )

        saved_name = f"{timestamp}_{original}"

        file.save(
            UPLOAD_DIR / saved_name
        )

    conn = get_db()

    conn.execute(
        """
        INSERT INTO faculty_applications
        (
            full_name,
            email,
            phone,
            role,
            cv_filename,
            note,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            data["full_name"].strip(),
            data["email"].strip(),
            data["phone"].strip(),
            data["role"].strip(),
            saved_name,
            data.get("note", "").strip(),
            datetime.now().isoformat(
                timespec="seconds"
            )
        )
    )

    conn.commit()
    conn.close()

    flash(
        "Your faculty profile has been submitted successfully.",
        "success"
    )

    return redirect(
        url_for("careers")
    )


@app.post("/contact")
def contact_submit():

    data = request.form

    if (
        not data.get("name", "").strip()
        or
        not data.get("email", "").strip()
        or
        not data.get("message", "").strip()
    ):

        flash(
            "Name, email and message are required.",
            "error"
        )

        return redirect(
            url_for("contact")
        )

    conn = get_db()

    conn.execute(
        """
        INSERT INTO contact_messages
        (
            name,
            email,
            phone,
            message,
            created_at
        )
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            data["name"].strip(),
            data["email"].strip(),
            data.get("phone", "").strip(),
            data["message"].strip(),
            datetime.now().isoformat(
                timespec="seconds"
            )
        )
    )

    conn.commit()
    conn.close()

    flash(
        "Message sent successfully.",
        "success"
    )

    return redirect(
        url_for("contact")
    )


@app.get("/api/courses")
def api_courses():
    return jsonify(COURSES)


@app.get("/health")
def health():
    return {
        "status": "ok",
        "service": "VIJAY Educational Services"
    }


if __name__ == "__main__":
    init_db()

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),
        debug=True
    )