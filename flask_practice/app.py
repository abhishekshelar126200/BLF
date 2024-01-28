from flask import Flask,render_template,request,redirect,session
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
import pandas as pd
import os
import sys
from docxtpl import DocxTemplate
import smtplib
import ssl
from email.message import EmailMessage
import string
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO
import base64
from sqlalchemy import LargeBinary


app=Flask(__name__)
app.secret_key = "login"
app.config['SQLALCHEMY_DATABASE_URI'] = "sqlite:///database.db"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)



class Todo(db.Model):
    SNO=db.Column(db.Integer,primary_key=True)
    title=db.Column(db.String(200),nullable=False)
    desc=db.Column(db.String(500),nullable=False)
    date_created=db.Column(db.DateTime,default=datetime.utcnow)
    excel = db.Column(db.String(200))
    word = db.Column(db.String(200))
    
    
class Database(db.Model):
    sno = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    desc = db.Column(db.String(500), nullable=False)
    

with app.app_context():
   db.create_all()

global df


@app.route("/admin",methods=['GET','POST'])
def index():
    # global dfe
    # global dfw
    global df
    if request.method=='POST':
        title=request.form['class']
        desc=request.form['branch']
        excel1=request.files['excel']
        word1=request.files['word']
        excel=excel1.filename
        word=word1.filename
        UPLOAD_FOLDER = f"D:/flask_practice/{title}/uploads"
        app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

        # If the user does not select a file, browser also sends an empty part without a filename
        if word1.filename == '':
            return "No selected file"

        if word1:
            # Save the file to the branch's subfolder inside the 'uploads' directory
            branch_upload_folder = os.path.join(app.config['UPLOAD_FOLDER'], desc)
            if not os.path.exists(branch_upload_folder):
                os.makedirs(branch_upload_folder)

            file_path = os.path.join(branch_upload_folder, word1.filename)
            word1.save(file_path)
        
        df=pd.read_excel(excel1)
        todo=Todo(title=title, desc=desc, excel=excel, word=word)

        db.session.add(todo)
        db.session.commit()

    return "Data Added Succesfully"


@app.route("/display")
def display():

    alltodos = Todo.query.all()
    return render_template('index.html',alltodos=alltodos)




@app.route("/delete/<int:SNO>")
def delete(SNO):

    todo=Todo.query.filter_by(SNO=SNO).first()
    # del dfe[SNO]
    # del dfw[SNO]
    db.session.delete(todo)
    db.session.commit()
    return redirect("http://127.0.0.1:5000/display")


@app.route("/notices/<int:SNO>")
def notices(SNO):
    # global dfe
    # global dfw
    global df

    todo=Todo.query.get(SNO)
    fname2=f"D:/flask_practice/{todo.title}/uploads/{todo.desc}/{todo.word}"
    # df=dfe[SNO]
    os.chdir(f"D:/flask_practice/{todo.title}/Notices/{todo.desc}")
    name = df["Name"].values
    id = df["roll_no"].values
    branch = df["branch"].values
    year = df["year"].values
    pending_fees = df["pending fees"].values

    zipped = zip(name, id, branch, year, pending_fees)
    file_name = []

    for a, b, c, d, e in zipped:
        if 50000 <= e < 100000:
            doc=DocxTemplate(fname2)

            context = {"student_name": a, "student_id": b, "branch": c, "year": d, "pending_fees": e}
            doc.render(context)
            doc.save('{}.docx'.format(f"{a}_{c}_{d}"))
            file_name.append(f"{a}_{c}_{d}")

    return redirect("/")


@app.route('/')
def hello_world():
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('email', None)
    return render_template("login.html")


@app.route('/login', methods=['POST', 'GET'])
def login():
    if request.method == 'POST':
        username = request.form["username"]
        password = request.form["password"]
        if username == "ganesh" and password == "1234":
            session['email'] = username
            return render_template("index.html")
        else:
            msg = "Invalid username/password"
            return render_template("login.html", msg=msg)
    else:
        return render_template("login.html")
    


if __name__ == "__main__":
    app.run(debug=True)