# Hostel Care AI — Red Theme Demo Prototype

A Flask + SQLite demo for hostel outpasses, student records, room details, announcements, complaints, fee status and gate verification.

## Features
- Role-based demo portals: Student, Warden and Security Guard.
- Student records imported at first run from `data/hostel_students_sample.xlsx`.
- Student profile, registration/name search, room/bed information and sample hostel fee status.
- Outpass flow: Student request → Parent approval link → Warden approval → Security QR verification → Exit/Return event.
- Parent approval link shown in student portal for demo. It is not automatically sent by email or WhatsApp.
- QR contains an opaque token, not student personal details. Guard can scan (scanner library is loaded from a CDN where available) or search manually by registration number / token.
- Complaints, warden resolution acknowledgement, announcements and manual attendance.
- Mobile-friendly red theme.
- Excel (.xlsx) and CSV student import via Warden dashboard.

## Run locally (Windows PowerShell)
1. Install Python 3.11+.
2. Extract this ZIP.
3. Open PowerShell in the `hostel_care_ai_website` folder (the folder containing `app.py` and `requirements.txt`).
4. Run:
   ```powershell
   python -m pip install -r requirements.txt
   python app.py
   ```
5. Open `http://127.0.0.1:5000`.

## Demo login
- Student: select a sample registration number on the login screen (e.g. `251FA1`)
- Warden username: `warden`
- Security Guard username: `guard`

This prototype has demo role selection and no password-based authentication. Do not use real student/parent data or deploy publicly without college approval, authentication, HTTPS, CSRF protection, privacy controls, backups and security review.

## Excel columns
The sample Excel sheet contains: `Reg_No`, `Student_Name`, `Branch`, `Year`, `Section`, `Gender`, `Block`, `Floor`, `Room_No`, `Bed_No`, `Parent_Name`, `Parent_Phone`, `Parent_Email`, `Hostel_Fee_Total`, `Fee_Paid`, `Fee_Pending`.

Warden dashboard accepts `.xlsx` and `.csv`; matching registration numbers are updated. Keep uploaded student data private and only use authorized records.
