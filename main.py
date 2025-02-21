import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from settings import Settings

clients = pd.read_excel('clients.xlsx')

server = smtplib.SMTP('smtp.gmail.com', port=587)
server.starttls()
server.login(Settings().EMAIL, Settings().PASSWORD)

for index, client in clients.iterrows():
    msg = MIMEMultipart()
    msg['Subject'] = 'E-mail de teste'
    msg['From'] = Settings().EMAIL
    msg['To'] = client['Email']
    message = f'Olá, {client['Nome']}!, você recebeu um e-mail.'
    msg.attach(MIMEText(message, 'plain'))
    server.sendmail(msg['From'], msg['To'], msg.as_string())

server.quit()
