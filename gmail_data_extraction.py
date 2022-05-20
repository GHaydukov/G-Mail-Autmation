import email
import imaplib
import os

from xlwt import Workbook

# Constants for the E-Mail credentials.
# Instead of writing the sensitive information directly to the script,
# we use environment variables to store the E-Mail credentials.
EMAIL_ADDRESS = os.environ.get("EMAIL_USER")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASS")


def read_email() -> None:
    """Reads the content of an E-Mail"""

    # Setting the encoding and adding the sheet.
    # We do this, because we want to save the needed information
    # from the email in an Excel file.
    workbook = Workbook(encoding="utf-8")
    table = workbook.add_sheet("data")

    # If we want to send an email we use - smtp.gmail.com,
    # If we want to read a received email we use - imap.gmail.com
    with imaplib.IMAP4_SSL("imap.gmail.com", 993) as imap:
        imap.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        imap.select("inbox")

        # data is in this case is a tuple and
        # contains a string and a list with the specified E-Mail, in our case UNSEEN
        data = imap.search(None, "UNSEEN")

        # Create the header of each column in the first row.
        table.write(0, 0, "From")
        table.write(0, 1, "Date")
        table.write(0, 2, "Subject")
        table.write(0, 3, "E-Mail Content")

        # Checks data to see if there are actually unread E-Mails.
        # The b notation in front of strings is used to specify a bytes string in Python.
        if data[1][0] != b"":
            # We want to get the mail id's from the data tuple
            # and because the list with the id's have the second position in the list (index 1)
            # we specify it with data[1] and therefore create a list mail_ids.
            mail_ids = data[1]

            # But because all ids are just one record in the mail_ids variable,
            # we create one more list, so every single id has its own position in the list.
            # Since there is only one record in the mail_ids list, and apparently it's a string,
            # we specify the index - 0.
            id_list = mail_ids[0].split()

            # The logic here with the first and last ids are that we want to iterate through them
            # and read the content from every single one of them.
            first_id = int(id_list[0])
            last_id = int(id_list[-1])

            line = 1

            # i is the E-Mail id.
            for i in range(last_id, first_id - 1, -1):
                # Reads the E-Mail in Gmail.
                # After the line below the E-Mail in Gmail is marked as "read"
                data = imap.fetch(str(i), '(RFC822)')

                for response in data:

                    if isinstance(response, list):
                        # response is a list with tuples, we want to access the first tuple,
                        # so we write [0] and then we want to access the bytes with the
                        # needed information (delivered-to, received-by etc) in this tuple,
                        # so we write [1]. message is in this case of the type Message class.
                        message = email.message_from_string(str(response[0][1], "utf-8"))
                        email_from = message['from']
                        email_date = message['date']
                        email_subject = message['subject']

                        # Writes the date, the subject and from whom is the email received.
                        table.write(line, 0, email_from)
                        table.write(line, 1, email_date)
                        table.write(line, 2, email_subject)

                        # If the E-Mail is multipart.
                        # Multipart is the type of the structure of the E-Mail.
                        if message.is_multipart():

                            # Iterate over the parts of the E-Mail.
                            for part in message.walk():
                                content_type = part.get_content_type()

                                try:
                                    # Gets the body of the E-Mail.
                                    body = part.get_payload()

                                # TypeError is an exception in Python programming language
                                # that occurs when the data type of objects in an operation is inappropriate.
                                # For example, If you attempt to divide an integer with a string,
                                # the data types of the integer and the string object will not be compatible.
                                except TypeError:
                                    print("Inappropriate data type of objects!")

                                if content_type == "text/plain":
                                    # Writes text/plain E-Mails into an Excel file.
                                    table.write(line, 3, body)

                        # If the E-Mail consists only of plain text.
                        else:
                            # extract content type of email
                            content_type = message.get_content_type()

                            # get the email body
                            body = message.get_payload()
                            if content_type == "text/plain":
                                # Writes the content into the Excel file.
                                table.write(0, 3, body)

                line += 1
        else:
            print("There are no unread E-Mails!")

    # Finally, save the Excel file.
    workbook.save("Unread Email Info.xls")


read_email()
