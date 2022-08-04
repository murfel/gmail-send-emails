# https://developers.google.com/gmail/api/quickstart/python

# Do before first run: download OAuth client -> rename to credentials.json
# https://console.cloud.google.com/apis/credentials?project=hse-org

# https://developers.google.com/gmail/api/reference/rest

import base64
import json
import os.path
from email.message import EmailMessage

import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.json.
# https://developers.google.com/identity/protocols/oauth2/scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/gmail.insert',
          'https://www.googleapis.com/auth/gmail.modify',
          'https://www.googleapis.com/auth/gmail.send']


def get_creds():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def send(creds, from_email, to_email, subject, content, send=False):
    """Create and insert a draft email.
       Print the returned draft's message and id.
       Returns: Draft object, including draft id and message meta data.
    """
    # create gmail api client
    service = build('gmail', 'v1', credentials=creds)

    message = EmailMessage()

    message.set_content(content)

    message['To'] = to_email
    message['From'] = from_email
    message['Subject'] = subject

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {
            'raw': encoded_message
    }

    if send:
        result = service.users().messages().send(userId='me', body=create_message).execute()
        print(to_email, result)
    else:
        print(to_email, content)


def clean_yo(filename):
    with open(filename, 'r') as file:
        content = file.read()
    with open(filename, 'w') as file:
        file.write(content.replace('ё', 'е'))


def match_reviews(from_students_filename, from_mentors_filename):
    s = pd.read_csv(from_students_filename)
    m = pd.read_csv(from_mentors_filename)
    student_to_email = dict()
    mentor_to_email = dict()
    for _, row in s.iterrows():
        student_to_email[row['student']] = row['student_email']
    for _, row in m.iterrows():
        mentor_to_email[row['mentor']] = row['mentor_email']

    num_mentors = 22
    num_students = num_mentors * 3 - 4
    print(f'Reviews from students: {len(student_to_email)} / {num_students} (missing {num_students - len(student_to_email)})')
    print(f'Reviews from mentors: {len(mentor_to_email)} / {num_mentors} ({num_mentors - len(mentor_to_email)}), '
          f'{m.shape[0]} / {num_students} ({num_students - m.shape[0]})')

    students_without_mentor_review = set(s['student'].to_list()) - set(m['student'].to_list())
    mentors_without_all_student_reviews = set(m['mentor'].to_list()) - set(s['mentor'].to_list())

    print('One party submitted a review, the other party didn\'t:')
    print('Students without mentor review', students_without_mentor_review)
    print('Mentors without student review', mentors_without_all_student_reviews)

    s.drop(s[s['student'].isin(students_without_mentor_review)].index, inplace=True)
    m.drop(m[m['student'].isin(mentors_without_all_student_reviews)].index, inplace=True)

    student_to_email.update(mentor_to_email)
    return s, m, student_to_email
    # print(m.shape[0])
    # m.drop(m[    s[(s['mentor'] == m['mentor']) & (s['student'] == m['student'])].shape[0] > 0    ].index, inplace=True)
    # print(m.shape[0])
    # return [()]


def construct_messages_to(df, name_to_email, to_students):
    """
    header_from_students = ['Timestamp', 'student_email', 'student', 'mentor', 'good', 'improve', 'score', 'additional', 'additional-private']
    header_from_mentsors = ['Timestamp', 'mentor_email', 'mentor', 'student', 'score', 'percentage', 'good', 'improve', 'additional', 'additional-private']
    """
    to_whom = 'student' if to_students else 'mentor'

    prefix = 'Привет! Присылаем подзадержавшиеся отзывы. Спасибо за участие в проектах!\n'
    messages = []
    for _, row in df.iterrows():
        fields = """Студент:
        Ментор:
        Что было хорошо:
        Что можно улучшить:
        Дополнительные комментарии:
        Оценка:""".split('\n')
        if row[to_whom] not in name_to_email:
            print(row[to_whom], ' - email not found, skipping')
            continue
        to_email = name_to_email[row[to_whom]]
        content = prefix
        pairs = zip(fields, [row['student'], row['mentor'], row['good'], row['improve'], row['additional'], str(row['score'])])
        for field, value in pairs:
            content += f'{field.strip()} {value}\n'

        messages.append((to_email, content))

    return messages


def main():
    creds = get_creds()
    from_email = 'murnatty@gmail.com'
    subject = '[Проекты C++] Финальный отзыв 2022'

    from_students_filename = 'from_students.csv'
    from_mentors_filename = 'from_mentors.csv'
    clean_yo(from_students_filename)
    clean_yo(from_mentors_filename)
    s, m, name_to_email = match_reviews(from_students_filename, from_mentors_filename)

    messages_to_mentors = construct_messages_to(s, name_to_email, to_students=False)
    # for msg in messages_to_mentors:
    #     to_email, content = msg
    #     print(from_email, to_email)
    #     send(creds, from_email, to_email, subject, content, send=False)


    # print('===================== От менторов студеетам =======================')
    messages_to_students = construct_messages_to(m, name_to_email, to_students=True)
    for msg in messages_to_students:
        to_email, content = msg
        print(from_email, to_email)
        send(creds, from_email, to_email, subject, content, send=False)


if __name__ == '__main__':
    main()
