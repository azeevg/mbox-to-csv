import mailbox
import csv
import html2text

# get body of email
def get_message(message):
    message_type = None
    if message.is_multipart():
        body = ""
        for part in message.get_payload():
            if part.is_multipart():
                for subpart in part.walk():
                    if subpart.get_content_type() == "text/html":
                        return subpart._payload, None
                        message_type = 'text/html' 
                        body += str(subpart)
                    elif subpart.get_content_type() == "text/plain":
                        return subpart._payload, None
                        message_type = 'text/plain' 
                        body += str(subpart)
            else:
                return part._payload, None
                message_type = part.get_content_type()
                body += str(part)
        return body, message_type
    else:
        return message.get_payload(), message_type

import sys

if __name__ == "__main__":

    # get mbox file
    #mbox_file = raw_input("name of mbox file in current directory (ex. my_file.mbox): ")
    mbox_file = sys.argv[1]
    # get name to filter
    #name_filter = raw_input("name of sender that you want to filter (ex. Jarrod Parkes, * for any): ")
    name_filter = ''
    # get email to filter
    #email_filter = raw_input("email of sender that you want to filter (ex. parkesfjarrod@gmail.com, * for any): ")
    email_filter = ''
    # create CSV file
    export_file_name = mbox_file + "_payload_0.csv"
    writer = csv.writer(open(export_file_name, "wb"))

    # create header row
    writer.writerow(["subject", "from", "date", "body"])

    print 'Start processing'
    # add rows based on mbox file
    for message in mailbox.mbox(mbox_file):
        contents, message_type = get_message(message)
        contents = html2text.comment_cleaner.sub('', contents)
        #if message_type and message_type == 'text/html':
        contents = html2text.html2text(contents)
        if contents:
            contents = contents.replace('\n', ' ').encode('utf-8', errors='ignore')
        # does message contain name or email filter?
        if name_filter != "" and name_filter in message["from"]:
            writer.writerow([message["subject"], message["from"], message["date"], contents])
        elif email_filter != "" and email_filter in message["from"]:
            writer.writerow([message["subject"], message["from"], message["date"], contents])
        elif email_filter != "" and name_filter != "":
            writer.writerow([message["subject"], message["from"], message["date"], contents])
        else:
            writer.writerow([message["subject"], message["from"], message["date"], contents])
            

    # print finish message
    print "generated csv file called " + export_file_name

