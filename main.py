import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class Email:

    def __init__(self, email, password, SMTP: str = None, port: int = None):

        self.email = email
        self.password = password

        if SMTP == None:
            domain = email.split("@")[1].lower()

            if domain == "gmail":
                self.smtp_server = "smtp.gmail.com"
            elif domain == "hotmail":
                self.smtp_server = "smtp.office365.com"
                
            elif domain == "hotmail":
                self.smtp_server = "smtp.live.com"
            else:
                raise Exception("unidentified domain, please try to pass the server smtp and port!")
        else:
            self.smtp_server = SMTP
            
        if port == None:
            self.port = 587
        else:
            self.port = port
            
    def send_file(self, to, body, subject, file_path, file_name: str = None):

        file_type = file_path.split("/")[-1].split(".")[1]

        if file_name == None:
            file_name = file_path.split("/")[-1]

        if file_name.find(".") == -1:
            file_name = file_name+"."+file_type


        # Email configuration
        from_email = self.email
        to_email = to

        # Create the message
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Add file to email
        attachment = MIMEApplication(open(file_path, 'rb').read(), _subtype='xlsx')
        attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
        msg.attach(attachment)

        # send the email
        smtp = smtplib.SMTP(self.smtp_server, self.port)
        smtp.starttls()
        smtp.login(self.email, self.password)
        smtp.sendmail(from_email, to_email, msg.as_string())
        smtp.quit()
        

    def send_text(self, to, body, subject):

        # Cria a mensagem de e-mail
        msg = MIMEText(body)
        msg['From'] = self.email
        msg['To'] = to
        msg['Subject'] = subject

        # Envia o e-mail
        smtp = smtplib.SMTP(self.smtp_server, self.port)
        smtp.starttls()
        smtp.login(self.email, self.password)
        smtp.sendmail(self.email, to, msg.as_string())
        smtp.quit()

teste = Email("andrey.so@redeioa.com.br", "dovovgrlvuhzkjgo", "smtp.gmail.com")
teste.send_file("andreybroker02@gmail.com", "teste", "teste com arquivo", './teste.txt', "outro_nome")