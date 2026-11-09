from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort
from functools import wraps
from pathlib import Path
import sqlite3, secrets, csv, io, os, base64
from datetime import datetime
from openpyxl import load_workbook
import qrcode
from qrcode.image.svg import SvgImage

app = Flask(__name__)
app.secret_key = os.environ.get('HOSTEL_CARE_SECRET', 'demo-only-change-this-secret')
BASE = Path(__file__).resolve().parent
DB_PATH = BASE / 'instance' / 'hostel.db'
DB_PATH.parent.mkdir(exist_ok=True)


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys=ON')
    return conn


def init_db():
    conn = db()
    conn.executescript('''
    CREATE TABLE IF NOT EXISTS students(
      id INTEGER PRIMARY KEY AUTOINCREMENT, reg_no TEXT UNIQUE NOT NULL, name TEXT NOT NULL,
      branch TEXT DEFAULT '', year INTEGER DEFAULT 1, section TEXT DEFAULT '', gender TEXT DEFAULT '',
      block TEXT DEFAULT '', floor INTEGER DEFAULT 1, room_no TEXT DEFAULT '', bed_no INTEGER DEFAULT 1,
      parent_name TEXT DEFAULT '', parent_phone TEXT DEFAULT '', parent_email TEXT DEFAULT '',
      fee_total REAL DEFAULT 0, fee_paid REAL DEFAULT 0, fee_pending REAL DEFAULT 0,
      attendance TEXT DEFAULT 'Present', active INTEGER DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS outpasses(
      id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL, destination TEXT NOT NULL,
      reason TEXT NOT NULL, from_date TEXT NOT NULL, to_date TEXT NOT NULL,
      parent_status TEXT DEFAULT 'Pending', warden_status TEXT DEFAULT 'Locked',
      parent_token TEXT UNIQUE NOT NULL, qr_token TEXT UNIQUE NOT NULL,
      gate_status TEXT DEFAULT 'Not exited', parent_note TEXT DEFAULT '', warden_note TEXT DEFAULT '',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS complaints(
      id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL, issue TEXT NOT NULL,
      status TEXT DEFAULT 'Open', resolution TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      resolved_at TEXT DEFAULT '', FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS announcements(
      id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, body TEXT NOT NULL,
      posted_at TEXT DEFAULT CURRENT_TIMESTAMP, author TEXT DEFAULT 'Warden'
    );
    CREATE TABLE IF NOT EXISTS gate_events(
      id INTEGER PRIMARY KEY AUTOINCREMENT, outpass_id INTEGER NOT NULL, event TEXT NOT NULL,
      event_at TEXT DEFAULT CURRENT_TIMESTAMP, guard_username TEXT DEFAULT 'guard',
      FOREIGN KEY(outpass_id) REFERENCES outpasses(id) ON DELETE CASCADE
    );
    ''')
    conn.commit()
    # Import the provided sample Excel on first run. Values are sample/demo data.
    if conn.execute('SELECT COUNT(*) FROM students').fetchone()[0] == 0:
        seed_file = BASE / 'data' / 'hostel_students_sample.xlsx'
        if seed_file.exists():
            wb=load_workbook(seed_file, data_only=True)
            ws=wb.active
            headers=[str(c.value).strip() for c in ws[1]]
            for values in ws.iter_rows(min_row=2, values_only=True):
                row=dict(zip(headers, values))
                reg=str(row.get('Reg_No') or '').strip()
                name=str(row.get('Student_Name') or '').strip()
                if not reg or not name: continue
                conn.execute('''INSERT OR IGNORE INTO students(reg_no,name,branch,year,section,gender,block,floor,room_no,bed_no,parent_name,parent_phone,parent_email,fee_total,fee_paid,fee_pending)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(
                    reg,name,str(row.get('Branch') or ''),int(row.get('Year') or 1),str(row.get('Section') or ''),str(row.get('Gender') or ''),
                    str(row.get('Block') or ''),int(row.get('Floor') or 1),str(row.get('Room_No') or ''),int(row.get('Bed_No') or 1),
                    str(row.get('Parent_Name') or ''),str(row.get('Parent_Phone') or ''),str(row.get('Parent_Email') or ''),
                    float(row.get('Hostel_Fee_Total') or 0),float(row.get('Fee_Paid') or 0),float(row.get('Fee_Pending') or 0)))
    if conn.execute('SELECT COUNT(*) FROM announcements').fetchone()[0] == 0:
        conn.execute('INSERT INTO announcements(title,body) VALUES(?,?)',('Welcome to Hostel Care AI','Sample announcement: Please check the hostel notice board for updates.'))
    conn.commit(); conn.close()


def login_required(role=None):
    def deco(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            if not session.get('role'):
                return redirect(url_for('login'))
            if role and session.get('role') != role:
                flash('You do not have permission to access that page.')
                return redirect(url_for('dashboard'))
            return fn(*args, **kwargs)
        return wrapped
    return deco


def get_student(conn, sid):
    return conn.execute('SELECT * FROM students WHERE id=?', (sid,)).fetchone()

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        role=request.form.get('role','student')
        username=request.form.get('username','').strip()
        conn=db()
        if role=='student':
            st=conn.execute('SELECT * FROM students WHERE reg_no=? AND active=1',(username.upper(),)).fetchone()
            conn.close()
            if not st:
                flash('Registration number not found. Try one of the sample registrations listed below.')
                return redirect(url_for('login'))
            session.clear(); session.update(role='student', username=st['reg_no'], student_id=st['id'])
        elif role=='warden' and username.lower()=='warden':
            conn.close(); session.clear(); session.update(role='warden', username='warden')
        elif role=='guard' and username.lower()=='guard':
            conn.close(); session.clear(); session.update(role='guard', username='guard')
        else:
            conn.close(); flash('For demo access, use Warden/warden or Security Guard/guard. Student uses their registration number.')
            return redirect(url_for('login'))
        return redirect(url_for('dashboard'))
    conn=db(); sample=conn.execute('SELECT reg_no,name,year FROM students ORDER BY year,reg_no').fetchall(); conn.close()
    return render_template('login.html', sample=sample)

@app.route('/dashboard')
@login_required()
def dashboard():
    role=session['role']
    if role=='student': return redirect(url_for('student_dashboard'))
    if role=='warden': return redirect(url_for('warden_dashboard'))
    return redirect(url_for('guard_dashboard'))

@app.route('/student')
@login_required('student')
def student_dashboard():
    conn=db(); st=get_student(conn,session['student_id'])
    passes=conn.execute('SELECT * FROM outpasses WHERE student_id=? ORDER BY id DESC',(st['id'],)).fetchall()
    active=[p for p in passes if p['gate_status']!='Returned']
    announcements=conn.execute('SELECT * FROM announcements ORDER BY id DESC LIMIT 8').fetchall()
    complaints=conn.execute('SELECT * FROM complaints WHERE student_id=? ORDER BY id DESC',(st['id'],)).fetchall()
    conn.close()
    return render_template('dashboard.html',role='student',student=st,passes=passes,active_passes=active,announcements=announcements,complaints=complaints)

@app.route('/profile')
@login_required('student')
def profile():
    conn=db(); st=get_student(conn,session['student_id']); conn.close()
    return render_template('profile.html',student=st)

@app.route('/outpass',methods=['GET','POST'])
@login_required('student')
def outpass():
    if request.method=='POST':
        destination=request.form.get('destination','').strip(); reason=request.form.get('reason','').strip()
        start=request.form.get('from_date',''); end=request.form.get('to_date','')
        if not destination or not reason or not start or not end or end < start:
            flash('Please complete the form and ensure the return date is not before the departure date.')
            return redirect(url_for('outpass'))
        conn=db(); token=secrets.token_urlsafe(24); qr_token=secrets.token_urlsafe(24)
        conn.execute('INSERT INTO outpasses(student_id,destination,reason,from_date,to_date,parent_token,qr_token) VALUES(?,?,?,?,?,?,?)',
          (session['student_id'],destination,reason,start,end,token,qr_token))
        conn.commit(); conn.close(); flash('Outpass request submitted. Parent approval is required first.')
        return redirect(url_for('student_dashboard'))
    return render_template('outpass_form.html')

@app.route('/parent/<token>',methods=['GET','POST'])
def parent_approval(token):
    conn=db(); p=conn.execute('SELECT o.*,s.name,s.reg_no FROM outpasses o JOIN students s ON s.id=o.student_id WHERE parent_token=?',(token,)).fetchone()
    if not p: conn.close(); return 'Approval link not found.',404
    if request.method=='POST' and p['parent_status']=='Pending':
        decision=request.form.get('decision')
        if decision in ('Approved','Rejected'):
            conn.execute('UPDATE outpasses SET parent_status=?,warden_status=? WHERE id=?',(decision,'Pending' if decision=='Approved' else 'Locked',p['id']))
            conn.commit(); p=conn.execute('SELECT o.*,s.name,s.reg_no FROM outpasses o JOIN students s ON s.id=o.student_id WHERE o.id=?',(p['id'],)).fetchone()
    conn.close(); return render_template('parent.html',p=p)

@app.route('/warden')
@login_required('warden')
def warden_dashboard():
    conn=db()
    ready=conn.execute("SELECT o.*,s.name,s.reg_no FROM outpasses o JOIN students s ON s.id=o.student_id WHERE o.parent_status='Approved' AND o.warden_status='Pending' ORDER BY o.id DESC").fetchall()
    students=conn.execute('SELECT * FROM students ORDER BY year,reg_no').fetchall()
    allpasses=conn.execute('SELECT o.*,s.name,s.reg_no FROM outpasses o JOIN students s ON s.id=o.student_id ORDER BY o.id DESC').fetchall()
    complaints=conn.execute("SELECT c.*,s.name,s.reg_no FROM complaints c JOIN students s ON s.id=c.student_id WHERE c.status='Open' ORDER BY c.id DESC").fetchall()
    announcements=conn.execute('SELECT * FROM announcements ORDER BY id DESC').fetchall()
    total=conn.execute('SELECT COUNT(*) FROM students').fetchone()[0]
    present=conn.execute("SELECT COUNT(*) FROM students WHERE attendance='Present'").fetchone()[0]
    fee=conn.execute('SELECT COALESCE(SUM(fee_pending),0) FROM students').fetchone()[0]
    conn.close()
    return render_template('dashboard.html',role='warden',ready=ready,students=students,allpasses=allpasses,complaints=complaints,announcements=announcements,total=total,present=present,fee=fee)

@app.route('/warden/student-search')
@login_required('warden')
def student_search():
    q=request.args.get('q','').strip()
    conn=db(); results=[]
    if q: results=conn.execute('SELECT * FROM students WHERE reg_no LIKE ? OR name LIKE ? ORDER BY reg_no LIMIT 50',(f'%{q}%',f'%{q}%')).fetchall()
    conn.close(); return render_template('search.html',q=q,results=results)

@app.route('/warden/approve/<int:pid>',methods=['POST'])
@login_required('warden')
def warden_approve(pid):
    conn=db(); p=conn.execute('SELECT parent_status FROM outpasses WHERE id=?',(pid,)).fetchone()
    if p and p['parent_status']=='Approved': conn.execute("UPDATE outpasses SET warden_status='Approved' WHERE id=?",(pid,)); flash('Outpass approved. QR verification is now available.')
    else: flash('Warden approval is locked until parent approval.')
    conn.commit(); conn.close(); return redirect(url_for('warden_dashboard'))

@app.route('/warden/reject/<int:pid>',methods=['POST'])
@login_required('warden')
def warden_reject(pid):
    conn=db(); conn.execute("UPDATE outpasses SET warden_status='Rejected' WHERE id=? AND parent_status='Approved'",(pid,)); conn.commit(); conn.close(); flash('Outpass rejected.'); return redirect(url_for('warden_dashboard'))

@app.route('/warden/attendance/<int:sid>',methods=['POST'])
@login_required('warden')
def attendance(sid):
    value=request.form.get('attendance')
    if value not in ('Present','Absent'): abort(400)
    conn=db(); conn.execute('UPDATE students SET attendance=? WHERE id=?',(value,sid)); conn.commit(); conn.close(); flash('Attendance updated.'); return redirect(request.referrer or url_for('warden_dashboard'))

@app.route('/warden/announcement',methods=['POST'])
@login_required('warden')
def announcement():
    title=request.form.get('title','').strip(); body=request.form.get('body','').strip()
    if title and body:
        conn=db(); conn.execute('INSERT INTO announcements(title,body,author) VALUES(?,?,?)',(title,body,'Warden')); conn.commit(); conn.close(); flash('Announcement posted.')
    return redirect(url_for('warden_dashboard'))

@app.route('/complaint',methods=['POST'])
@login_required('student')
def complaint_submit():
    issue=request.form.get('issue','').strip()
    if issue:
        conn=db(); conn.execute('INSERT INTO complaints(student_id,issue) VALUES(?,?)',(session['student_id'],issue)); conn.commit(); conn.close(); flash('Complaint submitted to the warden.')
    return redirect(url_for('student_dashboard'))

@app.route('/warden/complaint/<int:cid>/resolve',methods=['POST'])
@login_required('warden')
def complaint_resolve(cid):
    resolution=request.form.get('resolution','Problem solved.').strip() or 'Problem solved.'
    conn=db(); conn.execute("UPDATE complaints SET status='Resolved',resolution=?,resolved_at=CURRENT_TIMESTAMP WHERE id=?",(resolution,cid)); conn.commit(); conn.close(); flash('Complaint marked resolved. The student will see the acknowledgement.')
    return redirect(url_for('warden_dashboard'))

@app.route('/rooms')
@login_required()
def rooms():
    conn=db(); rows=conn.execute('SELECT * FROM students ORDER BY block,floor,room_no,bed_no').fetchall(); conn.close()
    return render_template('rooms.html',rows=rows,role=session['role'])

@app.route('/fees')
@login_required('student')
def fees():
    conn=db(); st=get_student(conn,session['student_id']); conn.close()
    return render_template('fees.html',student=st)

@app.route('/guard')
@login_required('guard')
def guard_dashboard():
    conn=db(); passes=conn.execute("SELECT o.*,s.name,s.reg_no FROM outpasses o JOIN students s ON s.id=o.student_id WHERE o.warden_status='Approved' ORDER BY o.id DESC").fetchall(); conn.close()
    return render_template('guard.html',passes=passes)

@app.route('/guard/verify',methods=['GET','POST'])
@login_required('guard')
def guard_verify():
    q=request.values.get('q','').strip()
    conn=db(); p=None
    if q:
        p=conn.execute("SELECT o.*,s.name,s.reg_no,s.branch,s.year FROM outpasses o JOIN students s ON s.id=o.student_id WHERE o.qr_token=? OR s.reg_no=? OR CAST(o.id AS TEXT)=? ORDER BY o.id DESC LIMIT 1",(q,q,q)).fetchone()
    conn.close()
    return render_template('guard_verify.html',p=p,q=q)

@app.route('/guard/pass/<int:pid>/event',methods=['POST'])
@login_required('guard')
def gate_event(pid):
    event=request.form.get('event')
    if event not in ('Exit','Return'): abort(400)
    conn=db(); p=conn.execute("SELECT * FROM outpasses WHERE id=? AND parent_status='Approved' AND warden_status='Approved'",(pid,)).fetchone()
    if not p: conn.close(); flash('This outpass is not fully approved.'); return redirect(url_for('guard_dashboard'))
    if event=='Exit':
        conn.execute("UPDATE outpasses SET gate_status='Exited' WHERE id=? AND gate_status IN ('Not exited','Returned')",(pid,))
    else:
        conn.execute("UPDATE outpasses SET gate_status='Returned' WHERE id=? AND gate_status='Exited'",(pid,))
    conn.execute('INSERT INTO gate_events(outpass_id,event,guard_username) VALUES(?,?,?)',(pid,event,session['username']))
    conn.commit(); conn.close(); flash(f'{event} recorded.'); return redirect(url_for('guard_verify',q=request.form.get('q','')))

@app.route('/qr/<int:pid>')
@login_required('student')
def qr_pass(pid):
    conn=db(); p=conn.execute('SELECT * FROM outpasses WHERE id=? AND student_id=?',(pid,session['student_id'])).fetchone(); conn.close()
    if not p or p['parent_status']!='Approved' or p['warden_status']!='Approved': abort(404)
    # QR encodes the opaque verification token, not student details.
    img=qrcode.make(p['qr_token'])
    buf=io.BytesIO(); img.save(buf,format='PNG')
    return render_template('qr.html',p=p,qr=base64.b64encode(buf.getvalue()).decode())


@app.route('/warden/import-students', methods=['POST'])
@login_required('warden')
def import_students():
    file=request.files.get('file')
    if not file or not file.filename:
        flash('Choose an Excel or CSV file first.')
        return redirect(url_for('warden_dashboard'))
    filename=file.filename.lower()
    try:
        if filename.endswith('.xlsx'):
            wb=load_workbook(file, data_only=True, read_only=True)
            ws=wb.active
            headers=[str(c.value or '').strip() for c in ws[1]]
            rows=[dict(zip(headers, vals)) for vals in ws.iter_rows(min_row=2, values_only=True)]
        elif filename.endswith('.csv'):
            text=file.stream.read().decode('utf-8-sig')
            rows=list(csv.DictReader(io.StringIO(text)))
        else:
            flash('Please upload an .xlsx or .csv file.')
            return redirect(url_for('warden_dashboard'))
        def val(row, *keys):
            norm={str(k).strip().lower().replace(' ','_'):v for k,v in row.items()}
            for key in keys:
                value=norm.get(key.lower().replace(' ','_'))
                if value is not None and str(value).strip()!='': return value
            return ''
        conn=db(); count=0
        for row in rows:
            reg=str(val(row,'Reg_No','reg_no','registration_number')).strip()
            name=str(val(row,'Student_Name','name','student_name')).strip()
            if not reg or not name: continue
            year=int(float(val(row,'Year','year') or 1))
            total=float(val(row,'Hostel_Fee_Total','fee_total') or 0)
            paid=float(val(row,'Fee_Paid','fee_paid') or 0)
            pending_val=val(row,'Fee_Pending','fee_pending')
            pending=float(pending_val) if str(pending_val).strip() else max(total-paid,0)
            fields=(name,str(val(row,'Branch','branch')),year,str(val(row,'Section','section')),str(val(row,'Gender','gender')),
              str(val(row,'Block','block')),int(float(val(row,'Floor','floor') or 1)),str(val(row,'Room_No','room_no')),
              int(float(val(row,'Bed_No','bed_no') or 1)),str(val(row,'Parent_Name','parent_name')),
              str(val(row,'Parent_Phone','parent_phone')),str(val(row,'Parent_Email','parent_email')),total,paid,pending)
            sql = """INSERT INTO students(reg_no,name,branch,year,section,gender,block,floor,room_no,bed_no,parent_name,parent_phone,parent_email,fee_total,fee_paid,fee_pending)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(reg_no) DO UPDATE SET
                name=excluded.name,branch=excluded.branch,year=excluded.year,section=excluded.section,gender=excluded.gender,
                block=excluded.block,floor=excluded.floor,room_no=excluded.room_no,bed_no=excluded.bed_no,
                parent_name=excluded.parent_name,parent_phone=excluded.parent_phone,parent_email=excluded.parent_email,
                fee_total=excluded.fee_total,fee_paid=excluded.fee_paid,fee_pending=excluded.fee_pending,active=1"""
            conn.execute(sql, (reg,*fields))
            count+=1
        conn.commit(); conn.close()
        flash(f'Imported or updated {count} student records.')
    except Exception:
        flash('Import failed. Check the file format and required columns, then try again.')
    return redirect(url_for('warden_dashboard'))

@app.route('/chatbot',methods=['GET','POST'])
@login_required()
def chatbot():
    if request.method=='POST':
        q=request.form.get('message','').lower()
        answer='I can help with hostel outpasses, approvals, rooms, complaints, announcements and fees. Ask me about one of these topics.'
        if any(x in q for x in ['outpass','out pass','pass']): answer='Student applies for outpass → parent approves → warden gives final approval → security guard verifies QR/token and marks Exit/Return.'
        elif any(x in q for x in ['complaint','problem','issue']): answer='Use the Complaint Box on your student dashboard. The warden can resolve it and you will see the acknowledgement and status.'
        elif any(x in q for x in ['room','bed','block']): answer='Open Room & Bed View to see your assigned Block, Floor, Room and Bed. Contact the warden for allocation changes.'
        elif any(x in q for x in ['fee','payment','paid']): answer='Open Hostel Fee Status on your dashboard to see total fee, paid amount and pending amount.'
        elif any(x in q for x in ['guard','qr','gate','exit','return']): answer='After both parent and warden approve, open the QR Outpass card. The security guard can scan the QR or enter the verification token/registration number and record Exit or Return.'
        elif any(x in q for x in ['announcement','notice']): answer='Hostel announcements posted by the warden appear on your dashboard.'
        elif any(x in q for x in ['profile','details']): answer='Open My Student Profile to view your student and room details.'
        return jsonify({'answer':answer})
    return render_template('chatbot.html')

@app.route('/logout')
def logout():
    session.clear(); flash('You have been logged out.'); return redirect(url_for('home'))

init_db()
if __name__=='__main__':
    app.run(debug=True)
