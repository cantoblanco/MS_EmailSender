from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QVBoxLayout, QSplitter, QFileDialog
from PyQt5.QtCore import Qt
import sys
import requests
import json
import configparser
import os

class EmailSender(QWidget):
    def __init__(self):
        super().__init__()
        self.token = None
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        self.initUI()

        # Load default email content if exists
        try:
            if os.path.exists('email.html'):
                with open('email.html', 'r') as f:
                    self.content_input.setPlainText(f.read())
        except Exception as e:
            print(f"Error loading email.html: {e}")

    def initUI(self):
        # Left upper area: Recipient and Subject
        recipient_label = QLabel('收件人 (用逗号分隔):')
        self.recipient_input = QLineEdit()

        subject_label = QLabel('主题:')
        self.subject_input = QLineEdit()

        top_left_layout = QVBoxLayout()
        top_left_layout.addWidget(recipient_label)
        top_left_layout.addWidget(self.recipient_input)
        top_left_layout.addWidget(subject_label)
        top_left_layout.addWidget(self.subject_input)

        # Left lower area: Status and Error Messages
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)

        # Right area: Email Content
        content_label = QLabel('内容:')
        self.content_input = QTextEdit()

        # Import Button
        import_button = QPushButton('导入HTML')
        import_button.clicked.connect(self.import_html)

        # Send Button
        send_button = QPushButton('发送邮件')
        send_button.clicked.connect(self.send_email)

        # Horizontal Splitter
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.addLayout(top_left_layout)
        left_layout.addWidget(self.status_text)
        left_widget.setLayout(left_layout)

        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_layout.addWidget(content_label)
        right_layout.addWidget(self.content_input)
        right_layout.addWidget(import_button)  # Add import button here
        right_widget.setLayout(right_layout)

        hsplitter = QSplitter(Qt.Horizontal)
        hsplitter.addWidget(left_widget)
        hsplitter.addWidget(right_widget)

        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(hsplitter)
        main_layout.addWidget(send_button)

        self.setLayout(main_layout)
        self.setWindowTitle('白歌邮件发送器')

    def import_html(self):
        # Import HTML file content
        file_name, _ = QFileDialog.getOpenFileName(self, 'Open HTML File', '', 'HTML Files (*.html)')
        if file_name:
            try:
                with open(file_name, 'r') as f:
                    self.content_input.setPlainText(f.read())
            except Exception as e:
                print(f"Error reading file {file_name}: {e}")

    def authenticate(self):
        # Authentication logic to get the access token
        token_url = self.config['DEFAULT']['TOKEN_URL']
        client_id = self.config['DEFAULT']['CLIENT_ID']
        client_secret = self.config['DEFAULT']['CLIENT_SECRET']
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
        token_data = {
            "client_id": client_id,
            "scope": scope,
            "client_secret": client_secret,
            "grant_type": grant_type
        }

        token_r = requests.post(token_url, data=token_data)
        self.token = token_r.json().get("access_token")

    def send_email(self):
        # Email sending logic
        if not self.token:
            self.authenticate()
        
        recipients = self.recipient_input.text().split(',')
        subject = self.subject_input.text()
        content = self.content_input.toPlainText()

        # Validate inputs
        if not recipients or not subject or not content:
            print("All fields must be filled")
            return

        for recipient in recipients:
            recipient = recipient.strip()
            if not recipient or '@' not in recipient:
                print(f"Invalid email address: {recipient}")
                return
            recipient = recipient.strip()
            self.update_status(f"正在向 {recipient} 发送邮件...")
            email_url =  self.config['DEFAULT']['EMAIL_URL']
            email_headers = {
                "Authorization": "Bearer " + self.token,
                "Content-Type": "application/json"
            }
            email_data = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": content
                    },
                    "toRecipients": [{"emailAddress": {"address": recipient}}]
                }
            }

            response = requests.post(email_url, headers=email_headers, data=json.dumps(email_data))
            if response.status_code == 202:
                self.update_status(f"向 {recipient} 发送邮件成功")
            else:
                self.update_status(f"向 {recipient} 发送邮件失败: {response.json()}")

    def update_status(self, message):
        self.status_text.append(message)

def main():
    app = QApplication(sys.argv)
    ex = EmailSender()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
