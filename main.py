# -*- coding: utf-8 -*-
import json
import mysql.connector as db
import smtplib
import datetime
import xlwt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


try:
    jsonTxt = open("data.json")
except:
    print("error while reading json file")
    exit(1)


json = json.load(jsonTxt)
# sql
database = json["Database"]

# colors
colors = ["bgcolor=#eeeeee", "bgcolor=#ffffff"]

# mail
mail = json["Mail"]

# static texts
filename = "raport.xls"
first_text = "nagłówek"
second_text = "dodatkowa treść"
third_text = "dodatkowa treść pod tabelką"


# create xls file
def sql_xls(sql_handler, sql_result):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('Sheet 1')
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'dd/MM/yyyy HH:mm:ss'

# x and y are coords of cell
    x = 0
    y = 0

# names of rows
    for item in sql_handler:
        worksheet.write(x, y, item)
        y += 1

    x = 1
    y = 0

# data
    for item in sql_result:
        for i in item:
            if type(i) is str:
                worksheet.write(x, y, str(i))
            elif type(i) is datetime.datetime:
                worksheet.write(x, y, i, date_format)
            else:
                worksheet.write(x, y, i)
            y += 1
        x += 1
        y = 0

    workbook.save("temp.xls")


# convert query to html table
def query_to_table(query):
    connection = db.connect(host=database["host"], port=database["port"], user=database["user"],
                            passwd=database["password"], db=database["db"])

    dbhandler = connection.cursor()
    column_names = None
    result = None

    for execu in dbhandler.execute(query, multi=True):
        if execu.with_rows:
            result = execu.fetchall()
            column_names = dbhandler.column_names

    number_colors = len(colors)
    index = 0
    sql_xls(column_names, result)

    # first row - names of columns
    body = "<Table> <tr %s >" % (colors[index % number_colors])
    index += 1
    for item in column_names:
        body += "<td class='names'> %s </td>" % item

    body += "</tr>"

    # mid
    for item in result:
        bold = 0
        body += "<tr %s>" % (colors[index % number_colors])
        for i in item[0:-1]:
            if "SUMA" in str(i):
                bold = 1
            if bold == 0:
                body += "<td>%s</td>" % (str(i))
            else:
                body += "<td><b>%s</b></td>" % (str(i))
        body += "<td class='last_column'> %s</td> </tr>" % (str(item[-1]))
        index += 1

    # end of table
    body += "</Table>"
    return body


def send_email(email_content, send_to, subject):
    msg = MIMEMultipart()
    msg['From'] = mail["sender"]
    msg['To'] = ','.join(send_to)
    msg['Subject'] = subject

    body = email_content
    msg.attach(MIMEText(body, 'html', 'utf-8'))
    attachment = open("temp.xls", "rb")

    p = MIMEBase('application', 'octet-stream')

# add attachment to mail
    p.set_payload(attachment.read())

    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    p.add_header('Content-Type', 'text/html; charset=utf-8')

# add message to mail
    msg.attach(p)
    msg.add_header('Content-Type', 'text/html; charset=utf-8')
    s = smtplib.SMTP(mail["server_smtp"])
    s.starttls()
    s.login(mail["login"], mail["password"])
    text = msg.as_string()
    s.sendmail(mail["login"], send_to, text)

    s.quit()
    for send in send_to:
        print("successfully sent the mail to %s" % send)


def run(querry, send_to, subject):
    mes = """
    <head>
        <meta charset=' utf-8 '>
        <style type="text/css">
                table {
				text-align: center;

				border-collapse: collapse;
				margin-left:auto;Ł
					margin-right:auto;

			}

			td{
				border: 1px solid black;
				padding: 4px;
			}

			.last_column	{ text-align: right;}
			.names 		{ font-weight: bold; }
			.text_1 	{text-align: center;}
			.text_2		{text-align: left;}
			.text_3		{text-align: left;}
			.empty 		{border: 0px solid black}
			</style>
	 </head>
	"""

    mes += """
	<body>
		<h2 class='text_1'> %s 	</h2>
		<p class='text_2'> %s	</p>
		%s   <!-- tutaj jest tabela -->
		<p class ='text_3'> %s	</p>
	</body>

	""" % (first_text, second_text, query_to_table(querry), third_text)

    send_email(mes, send_to, subject)


if __name__ == "__main__":
    try:
        now = datetime.datetime.now()
        now = now + datetime.timedelta(days=-1)
        query = "Create temporary table temp select * from `DOCS_S1`; select * from temp;"
        send_to = ['Tomasz.kurek@hotmail.com']
        subject = u"ąśRaport z dnia %s.%s.%s" % (str(now.day), str(now.month), str(now.year))
        run(query, send_to, subject)
    except:
        print("failure sending mail")

    try:
        now = datetime.datetime.now()
        Querry = "SELECT* FROM `DOCS_S1`"
        send_to = ['testerraportow@gmail.com', 'Tomasz.kurek@hotmail.com']
        subject = u'Raport z dnia %s.%s.%s' % (str(now.day), str(now.month), str(now.year))
        run(Querry, send_to, subject)
    except:
        print("failure sending mail")


