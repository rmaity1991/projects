import smtplib


connection=smtplib.SMTP('smtp.gmail.com',587)

connection.ehlo()

connection.starttls()

connection.login('official.rohitmaity@gmail.com','41916082rohitmaity')

connection.sendmail(from_addr='official.rohitmaity@gmail.com',to_addr='official.rohitmaity@icloud.com',msg="This is a test email")

connection.quit()