from __future__ import print_function
import base64
from email.mime.text import MIMEText
import os
import pprint
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from apiclient import errors
from absl import app
from absl import flags
from absl import logging

logging.set_verbosity(logging.ERROR)

FLAGS = flags.FLAGS

# Flag names are globally defined!  So in general, we need to be
# careful to pick names that are unlikely to be used by other libraries.
# If there is a conflict, we'll get an error at import time.
flags.DEFINE_string("sender", None, "The email address to send messages from.")
flags.DEFINE_integer("tut", None, "Your age in years.", lower_bound=0)
flags.DEFINE_string(
    "recipient", None, "The email address to send to (for testing purposes)"
)
flags.DEFINE_string("spreadsheet_id", None, "The ID of the spreadsheet to check")
# TODO Logging level as flag

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]

# The ID and range of a sample spreadsheet.
RANGE_NAME = "metadata!A2:F"

pp = pprint.PrettyPrinter(indent=2)


def post_process(outer_array):
    pad_length = 6
    for inner_array in outer_array:
        for i in range(len(inner_array)):
            inner_array[i] = inner_array[i].strip()
        while len(inner_array) < pad_length:
            inner_array.append("")
    return outer_array


def create_draft(service, user_id, message_body):
    """Create and insert a draft email. Print the returned draft's message and id.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message_body: The body of the email message, including headers.

    Returns:
        Draft object, including draft id and message meta data.
    """
    try:
        message = {"message": message_body}
        draft = service.users().drafts().create(userId=user_id, body=message).execute()

        print("Draft id: {}\nDraft message: {}".format(draft["id"], draft["message"]))

        return draft
    except errors.HttpError as error:
        print("An error occurred: {}".format(error))
        return None


def create_message(sender, to, subject, message_text):
    """Create a message for an email.

    Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.

    Returns:
        An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject
    return {"raw": base64.urlsafe_b64encode(message.as_string())}


def send_message(service, user_id, message):
    """Send an email message.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.

    Returns:
        Sent Message.
    """
    try:
        message = (
            service.users().messages().send(userId=user_id, body=message).execute()
        )
        print("Message Id: %s" % message["id"])
        return message
    except errors.HttpError as error:
        print("An error occurred: %s" % error)


def main(_):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    sheet_service = build("sheets", "v4", cache_discovery=False, credentials=creds)
    gmail_service = build("gmail", "v1", cache_discovery=False, credentials=creds)

    tutorial_idx = FLAGS.tut + 2

    # Call the Sheets API
    sheet = sheet_service.spreadsheets()
    # result is a dict constructed from the json returned from the get call
    result = (
        sheet.values()
        .get(spreadsheetId=FLAGS.spreadsheet_id, range=RANGE_NAME)
        .execute()
    )
    # values is the value in the dict result for the key "values"
    # this get is the dict.get, *not* related to the get API
    values = post_process(result.get("values", []))

    if not values:
        print("No data found.")
    else:
        missing_tuts = []
        print(
            "Found the following tutors with missing tutorial {} sheet(s):".format(
                FLAGS.tut
            )
        )
        for row in values:
            if row[tutorial_idx] == "" and row[1] not in missing_tuts:
                missing_tuts.append(row[1])
        pp.pprint(missing_tuts)
        print("Send follow-up emails? y/N")
        send_messages = raw_input()
        if send_messages == "y":
            me = FLAGS.sender
            recipient = FLAGS.recipient
            message = create_message(me, recipient, "you late", "why hello")
            send_message(gmail_service, me, message)


if __name__ == "__main__":
    flags.mark_flag_as_required("sender")
    app.run(main)
