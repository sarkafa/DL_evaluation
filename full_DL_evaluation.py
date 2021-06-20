#generate pdf
import pyodbc
from jinja2 import Environment,FileSystemLoader
import xlrd
from xlrd import open_workbook, cellname
import datetime
import pandas as pd
import pdfkit
import os
import matplotlib.pyplot as plt
from matplotlib.pyplot import figure

#send email
import pyodbc
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os



#
##
### DB DATA DOWNLOAD
##
#

############## set parameters for sql server and table ##############
conn = pyodbc.connect('Driver={SQL Server};'
                      #hidden because of personnel data
                      #hidden because of personnel data
                      'Trusted_Connection=yes;')

#seelect all from table
cursor = conn.cursor().execute('SELECT *  FROM  #hidden table because of personnel data')

#set first row as columns names
columns = [column[0] for column in cursor.description]


############# read data from table to dict #############################
db_data = []
for row in cursor.fetchall():
    db_data.append(dict(zip(columns, row)))


#
##
### FILLING HTML TEMPLATE
##
#

########## generate chart ############################################

def generate_graph(values, response_key):
  
    total_score = values[0]
    detailed_score = values[1:]
    
    competences =[
        'E: Data interpretation, communication and decision-making',
        'D: Data analysis and evaluation',
        'C: Data collection and preparation', 
        'B: Analytical principles and methods',
        'A: Data concepts, ethiscs and protection'     
        ]
    
    fig, ax = plt.subplots();

    ax.spines['right'].set_color('#ebe3e3')
    ax.spines['top'].set_visible(False);
    ax.spines['bottom'].set_visible(False);
    plt.xticks([10, 20, 30, 40, 50, 60, 70, 80, 90, 100]);
    ax.set_xlim([0,100]);

    ax.set_axisbelow(True)
    ax.xaxis.grid(color='#ebe3e3', )

    plt.vlines(total_score, ymin=-0.5, ymax=100, linestyles='dashed', colors='#00957d');

    plt.barh(competences, detailed_score, color ='black' );
   
    fig.set_size_inches(10, 3)
    plt.tight_layout()
    plt.draw()
    plt.savefig('graphs/'+str(response_key)+'.png',facecolor='white', dpi=100, transparent=False);
    #plt.show();
    


########## fill html template with data from DB ######################
file_loader = FileSystemLoader(searchpath="./templates")
env = Environment(loader=file_loader)
template = env.get_template("template_likert_black_A4.html")

#for each respondent data fill the template
for respondent in range(0,len(db_data)):

    total_score     = db_data[respondent].get('Total_Score')
    code            = db_data[respondent].get('Code') 
    email           = db_data[respondent].get('Email')
    a_score         = db_data[respondent].get('A_Section')
    b_score         = db_data[respondent].get('B_Section')
    c_score         = db_data[respondent].get('C_Section') 
    d_score         = db_data[respondent].get('D_Section')
    e_score         = db_data[respondent].get('E_Section')
    submit_time     = db_data[respondent].get('Submit_Date')  
    total_score_meaning = db_data[respondent].get('Total_Score_Meaning')  
    response_key    = db_data[respondent].get('Responses_Key')
    a_Sec_Meaning    = db_data[respondent].get('A_Sec_Meaning')
    b_Sec_Meaning    = db_data[respondent].get('B_Sec_Meaning')
    c_Sec_Meaning    = db_data[respondent].get('C_Sec_Meaning')
    d_Sec_Meaning    = db_data[respondent].get('D_Sec_Meaning')
    e_Sec_Meaning    = db_data[respondent].get('E_Sec_Meaning')
    a_Sec_Recomm    = db_data[respondent].get('A_Sec_Recomm')
    b_Sec_Recomm    = db_data[respondent].get('B_Sec_Recomm')
    c_Sec_Recomm    = db_data[respondent].get('C_Sec_Recomm')
    d_Sec_Recomm    = db_data[respondent].get('D_Sec_Recomm')
    e_Sec_Recomm    = db_data[respondent].get('E_Sec_Recomm')

    #create respondents folder
    os.mkdir(os.getcwd()+'\\respondents_files\\'+response_key)


    values = [total_score, e_score, d_score, c_score, b_score, a_score]  
    generate_graph(values, response_key)  


    html_filling = template.render(

       total_score = total_score,
        code = code, 
        email = email, 
        a_score = a_score, 
        b_score = b_score,
        c_score = c_score, 
        d_score = d_score,
        e_score = e_score,
        submit_time = submit_time,
        total_score_meaning = total_score_meaning,
        graph = response_key,
        a_Sec_Meaning = a_Sec_Meaning,
        b_Sec_Meaning = b_Sec_Meaning, 
        c_Sec_Meaning = c_Sec_Meaning,
        d_Sec_Meaning = d_Sec_Meaning, 
        e_Sec_Meaning = e_Sec_Meaning,
        a_Sec_Recomm = a_Sec_Recomm, 
        b_Sec_Recomm = b_Sec_Recomm,
        c_Sec_Recomm = c_Sec_Recomm, 
        d_Sec_Recomm = d_Sec_Recomm,
        e_Sec_Recomm = e_Sec_Recomm
    )

    html_to_fill      = open('filled_html/'+str(response_key)+'.html', 'w')
    html_to_fill.write(html_filling)
    html_to_fill.close()





#
##
### CONVERTING TO PDF ####################################
##
#

###################### convert to PDF ####################
config = pdfkit.configuration(wkhtmltopdf='C:\\Program Files (x86)\\wkhtmltopdf\\bin\\wkhtmltopdf.exe')

options = {    
    'page-size': 'A4',
    'dpi': 96,
    'encoding':'utf-8', 
    'page-size':'A4',
    'margin-top':'0cm',
    'margin-bottom':'0cm',
    'margin-left':'0cm',
    'margin-right':'0cm',
    'disable-smart-shrinking': '',
    'enable-local-file-access': None
    }

directory_filled_html =os.getcwd() + '\\filled_html\\'


for file in os.listdir(directory_filled_html):
  
  directory_pdfs = os.getcwd()+'\\respondents_files\\' + file.split('.', 1)[0]
  print(str(file.split('.', 1)[0]) +'.pdf')
  pdfkit.from_file(directory_filled_html+str(file),directory_pdfs+'\\DL_Result.pdf',configuration=config, options=options)


#
##
### TRUNCATE GRAPHS FILE AND FILLED HTML FILES ####################################
##
#


directory_graphs = os.getcwd() + '\\graphs'
for file in os.listdir(directory_graphs):
    os.remove(os.path.join(directory_graphs, file))



directory_filled_html = os.getcwd() + '\\filled_html'
for file in os.listdir(directory_filled_html):
    os.remove(os.path.join(directory_filled_html, file))





#
##
### EMAIL
##
#
#read html body for email
with open('email_body/email_body.html', 'r') as f:
    html_email_body = f.read()


###################### defina mail class ####################
class Mail:
    def __init__(self):
        self.mail_host = 'hidden personnel data'
        self.mail_user = 'hidden personnel data'        

    def send_email(self, to, subject, body, path, attach):
        """ set parameters to email """

        msg = MIMEMultipart() 
        msg['From'] = self.mail_user
        msg['To'] = to
        msg['Subject'] = subject
        msg['Cc'] = self.mail_user

        #set email body 
        email_body = MIMEText(body, 'html')
        msg.attach(email_body)

        #attach pdf
        pdfname = attach
        binary_pdf = open(path + pdfname, 'rb')
        payload = MIMEBase('application', 'octate-stream', Name=pdfname)
        payload.set_payload((binary_pdf).read())
        encoders.encode_base64(payload)
        payload.add_header('Content-Decomposition', 'attachment', filename=pdfname)
        msg.attach(payload)

        #send mesage
        with smtplib.SMTP(self.mail_host, 25) as mailserver:
            mailserver.ehlo()
            mailserver.starttls()
            mailserver.send_message(msg)
            mailserver.quit()

         
###################### define function for loading data do DB about sent emails#########
def load_data_to_db(response_key, email, time_stamp_email,pdf_path, email_check):
    conn = pyodbc.connect('Driver={SQL Server};'
                      #hidden because of personnel data
                      #hidden because of personnel data
                      'Trusted_Connection=yes;')

    cursor = conn.cursor()

    cursor.execute(f'''
                    INSERT INTO #hidden table because of personnel data                     

                    ([Responses_Key]
                    ,[Email]
                    ,[Sent_Time]
                    ,[Pdf_Path]
                    ,[Email_Check])

                VALUES ('{response_key}', '{email}', '{time_stamp_email}', '{pdf_path}',  '{email_check}')
                    ''')
    conn.commit()


###################### send email ####################################################
# mail = Mail()

for respondent in range(0,len(db_data)):

    email_respondent = db_data[respondent].get('Email') 
    response_key     = db_data[respondent].get('Responses_Key') 
    time_stamp_email = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    pdf_path         = os.getcwd()+'\\respondents_files\\' + str(response_key) +'\\DL_Result.pdf'

    #verify if email exits    
    try:
        mail.send_email(email_respondent, 'Data literacy result', html_email_body, 'respondents_files/'+str(response_key)+'/', 'DL_Result.pdf' )
    except smtplib.SMTPRecipientsRefused as e:
        print(e)
        print( 'Email does not exist.')    
        load_data_to_db(response_key, email_respondent, time_stamp_email, pdf_path, email_check=0)
    else:   
        load_data_to_db(response_key, email_respondent, time_stamp_email, pdf_path, email_check=1)
        
        