#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2021-11-29
Wichtelbot
Alexander J. Pfleger

Secret Santa Programme, that generates random pairings and notifies
all participants over email whom they have to gift something.

Sources:
    https://realpython.com/python-send-email/
"""

import smtplib
import ssl
from email.mime.text import MIMEText
import xlrd  # == 1.2.0 because they removed support for xlsx later
import random


picture = """
                                               ._...._.
                       \\ .*. /                .::o:::::.
                        (\\o/)                .:::'''':o:.
                         >*<                 `:)_>()<_(:'
                        >0<@<             @    `'//\\\\'`    @
                       >>>@<<*          @ #     //  \\\\     # @
                      >@>*<0<<<      .__#_#____/'____'\\____#_#_.
                     >*>>@<<<@<<     [_________________________]
                    >@>>0<<<*<<@<     |=_- .-/\\ /\\ /\\ /\\--. =_-|
                   >*>>0<<@<<<@<<<    |-_= | \\ \\ \\ \\ \\ \\\\-|-_=-|
                  >@>>*<<@<>*<<0<*<   |_=-=| / // // // / |_=-_|
    \\*/          >0>>*<<@<>0><<*<@<<  |=_- | `-'`-'`-'`-' |=_=-|
.___\\U//__.    >*>>@><0<<*>>@><*<0<<  | =_-| o          o |_==_|
 \\ | | \\  |  >@>>0<*<<0>>@<<0<<<*<@<  |=_- | !     (    ! |=-_=|
  \\| | _(UU)_ >((*))_>0><*<0><@<<<0<*<|-,-=| !    ).    ! |-_-=|
 \\ \\| || / //||.*.*.*.|>>@<<*<<@>><0<<|=_,=| ! __(:')__ ! |=_==|
 \\_|_|&&_// ||*.*.*.*|_\\db//_   (\\_/)-|     /^\\=^=^^=^=/^\\| _=_|
  ""|'.'.'.|~~|.*.*.*|      |  =('Y')=|=_,//.------------.\\\\_,_|
    |'.'.'.|  |^^^^^^|______|  ( ~~~ )|_,_/(((((((())))))))\\_,_|
    ~~~~~~~ ""       `------'  `w---w'|_____`------------'_____|
________________________________________________________________

"""


def create_message(nameA, nameB, mailB, date, adresse, food, comment):
    msg_str = f"Liebe(r) {nameA}!\n\n"
    msg_str += (
        f"Schön, dass Du dieses Jahr beim AFÖP-Wichteln dabei bist. Bitte bereite ein Geschenk für {nameB} vor aber "
        f"sags nicht weiter, wen Du beschenkst, damit alle mehr Freude am Wichteln haben. Um ein passendes Geschenk zu "
        f"finden und es auch rechtzeitig zu übergeben hat Dir Dein Wichtel noch ein paar Informationen hinterlassen."
        f"\n\n"
    )

    if len(adresse) != 0:
        msg_str += (
            f"Dein Wichtel würde das Geschenk am liebsten so erhalten:\n{adresse}\n\n"
        )

    if len(food) != 0:
        msg_str += f"Falls Du etwas Essbares schenken willst beachte bitte noch (Allergien etc.):\n{food}\n\n"

    if len(date) != 0:
        msg_str += (
            f"Voraussichtlich ist {nameB} noch bis {date} in Wien. Also beeile Dich, damit Dein Geschenk rechtzeitig "
            f"ankommt!\n\n"
        )

    msg_str += f"Im äußersten Notfall kannst {nameB} unter {mailB} erreichen, aber das könnte die Überraschung verderben.\n\n"

    if len(comment) != 0:
        msg_str += (
            f"Abschließen möchte Dir Dein Wichtel noch folgendes sagen:\n{comment}\n\n"
        )

    msg_str += "Ich wünsche Dir noch viel Spaß beim Schenken und Beschenkt werden.\n\tDein Wichtelbot\n\n"
    msg_str += picture

    return msg_str


exp_cols = 8

# read data from excel
excel_sheet = xlrd.open_workbook("participants.xlsx").sheet_by_index(0)

n_participants = excel_sheet.nrows

assert excel_sheet.ncols == exp_cols  # check if xlsx has proper dimension
assert n_participants > 2  # check if enough participants

# generate pairings
partA = list(range(1, n_participants))
partB = partA.copy()
# badpair = 1
# while badpair != 0:
#    badpair = 0
#    random.shuffle(partB)
#    for k in range(n_participants-1):
#        badpair += (partA[k] == partB[k])
while any(a == b for (a, b) in zip(partA, partB)):
    random.shuffle(partB)

# send mail
port = 465  # For SSL
smtp_server = "smtp.domainname.xyz"  # Enter your smtp_server
sender_email = "sender@mail.xyz"  # Enter your sender_email
password = input("Type your password and press enter: ")


for pair in range(n_participants - 1):
    receiver_email = excel_sheet.cell_value(partA[pair], 3)

    message = create_message(
        excel_sheet.cell_value(partA[pair], 2),
        excel_sheet.cell_value(partB[pair], 2),
        excel_sheet.cell_value(partB[pair], 3),
        excel_sheet.cell_value(partB[pair], 4),
        excel_sheet.cell_value(partB[pair], 5),
        excel_sheet.cell_value(partB[pair], 6),
        excel_sheet.cell_value(partB[pair], 7),
    )

    print(receiver_email)
    print(message)

    msg = MIMEText(message.encode("utf-8"), _charset="utf-8")
    msg["Subject"] = "Wichteln 2023"
    msg["From"] = sender_email
    msg["To"] = receiver_email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, msg.as_string())
