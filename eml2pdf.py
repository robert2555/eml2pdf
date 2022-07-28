import os
import argparse
import email
from email.header import decode_header
from fpdf import FPDF
import re
from warnings import filterwarnings


def writeErrorPdf(cause, filename):
    # Create a new PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', '', 8);

    # Write the "Empty" note in the pdf
    if cause == "empty":
        pdf.cell(190, 4, txt="EML does not contain text", align='C')

    # Write the "technical Error" note in the pdf
    if cause == "error":
        pdf.cell(190, 4,
                 txt="Due to special circumstances, there couldnt be created a valid Email-Text in this specific case.",
                 align='C')

    # Write PDF
    pdf.output(filename)


def writePdfFile(text, filename):
    # Check if text is not empty
    if not text:
        writeErrorPdf("empty", filename)
        return

    # pdf is a FPDF Object
    pdf = FPDF()

    # Add a page
    pdf.add_page()

    # set style and size of font
    pdf.add_font('DejaVu', '', '/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf', uni=True)
    pdf.set_font('DejaVu', '', 8)

    # Clear all Emoji symbols with regex
    text = re.sub(u'[\U0001f300-\U0001f64F]', "<EMOJI>", text)

    # insert the texts in pdf
    for line in text.split('\n'):
        pdf.multi_cell(190, 4, txt=line, align='L')

    # ignore warnings
    filterwarnings('ignore')

    # save the pdf with name .pdf
    try:
        pdf.output(filename)
    except:
        writeErrorPdf("error", filename)


def extractMailText(msg, filename):
    if msg.is_multipart():
        for v_str, v_charset in decode_header(msg['From']):
            # Decode it with charset
            if v_charset:
                # Try to encode with charset
                try:
                    v_from = v_str.decode(v_charset) + " "
                    continue
                except:
                    pass
                # Ignore bytes in string
                try:
                    v_from = v_str.decode("utf-8", "ignore") + " "
                    continue
                except:
                    pass
            # Try to decode it if its in bytes
            try:
                v_from = bytes.decode(v_str) + " "
                continue
            except:
                pass
            # If its a plain string
            v_from = v_str
        # Decode "To" field if encoded
        if decode_header(msg['To'])[0][1]:
            # Get string and charset
            dstring, charset = decode_header(msg['To'])[0]
            # Decode string with given charset
            v_to = dstring.decode(charset) + " "
        else:
            v_to = msg['To']
        # Decode "Subject" field if existing and encoded
        try:
            if decode_header(msg['Subject'])[0][1]:
                # Get string and charset
                dstring, charset = decode_header(msg['Subject'])[0]
                # Decode string with given charset
                v_subj = dstring.decode(charset) + " "
            else:
                v_subj = msg['Subject']
        except:
            v_subj = 'leer'

        # Join all header fields
        header = ['Von: ' + v_from, 'An: ' + v_to, 'Datum: ' + msg['Date'], 'Betreff: ' + v_subj]

        attach = "Anlagen: "
        body = []
        # Walk through the mail
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            if ctype == 'application/pdf' or ctype == 'application/octet-stream':
                # Decode filename if encoded
                if decode_header(part.get_filename())[0][1]:
                    # Get string and charset
                    dstring, charset = decode_header(part.get_filename())[0]
                    # Decode string with given charset
                    attach = attach + dstring.decode(charset) + " "
                else:
                    attach = attach + part.get_filename() + " "

            # Print a message for html based bodies
            if ctype == 'text/html':
                body.append("\n\n\n\nNOTICE: HTML-based Mail - text is not standardized\n")

            # Get the body
            if ctype == 'text/plain' or ctype == 'text/html':
                # Check for base64 encoded content
                try:
                    body.append(re.sub("([\{\<]).*?([\>\}])", "", part.get_payload(decode=True).decode()))
                    continue
                except:
                    pass
                # Get plain-text content and decode it with the given charset
                try:
                    body.append(re.sub("([\{\<]).*?([\>\}])", "",
                                       part.get_payload(decode=True).decode(part.get_content_charset())))
                    continue
                except:
                    pass
                # Get just the plain text
                try:
                    body.append(re.sub("([\{\<]).*?([\>\}])", "", part.get_payload()))
                    continue
                except:
                    pass
                # Continue
                continue
        # Create a list of lists
        header.append(attach)
        mail = header + ["\n"] + body
        # Return body list as string
        return '\n'.join(mail)


def main():
    # Add argparser
    parser = argparse.ArgumentParser()

    # Parse args
    parser.add_argument("-f", "--file", required=True, help="path value")
    args = parser.parse_args()
    filename = args.file

    # get mail content
    with open(filename, 'rb') as file:
        msg = email.message_from_bytes(file.read())

    # extract mail text
    body = extractMailText(msg, filename)

    # write PDF file
    writePdfFile(body, filename + ".pdf")


if __name__ == "__main__":
    main()

