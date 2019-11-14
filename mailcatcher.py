#!/usr/bin/env python3
#-*-coding:utf8;-*-
############################################################################
#   MailCatcher by Sven Gaechter Allround-Service.biz                      #
#                                                                          #
#   History:                                                               #
#   0.1: Initial quick and dirty Code                                      #
############################################################################

from imap_tools import MailBox, Q
import os
import logging
import configparser

##### Opening Config-File #####
config = configparser.ConfigParser()
config.read(os.path.join(os.path.dirname(__file__),'config.ini'))
login = config['General']['login']
password = config['General']['password']
senderAddr = config['General']['senderAddr']
attPath = config['General']['attPath']
mailServer = config['General']['mailServer']
destFilename = config['General']['destFilename']
fileExt = config['General']['fileExt']

##### Log-Handler #####
logging.basicConfig(filename=os.path.join(os.path.dirname(__file__),'Mailcatcher.log'),format='%(asctime)s - %(message)s', level=config['General']['loglevel'])
logging.info('-------------------------')
logging.info(f"{__file__} gestartet...")


mailbox = MailBox(mailServer)
mailbox.login(login, password, initial_folder='INBOX')
"""
try:
    mailbox.folder.create('INBOX/folder1')
except:
    print('Ordner existiert schon')
"""

for msg in mailbox.fetch(Q(from_=senderAddr)):
    Subject = msg.subject
    logging.debug(f'Betreff {Subject}')
    if len(msg.attachments) > 0:
        attachment = [i[0] for i in msg.attachments] + [i[1] for i in msg.attachments]

        if fileExt in attachment[0]:
            logging.info(f'{attachment[0]} wird heruntergeladen...')
            f = open(os.path.join(attPath, attachment[0]), "wb") #destFilename If the Filename will not be replaced
            f.write(attachment[1])
            f.close()
            try:
                mailbox.move(msg.uid, 'INBOX/processed')
            except:
                logging.debug(f'Mail mit dem Betreff {subject} konnte nicht verschoben werden')
        else:
            logging.info(f'{attachment[0]} entspricht nicht dem Suchbegriff {fileExt}. Datei wurde nicht heruntergeladen')
            try:
                mailbox.move(msg.uid, 'INBOX/rejected')
                continue
            except:
                logging.debug(f'Mail mit dem Betreff {subject} konnte nicht verschoben werden')
    else:
        logging.debug(f'Mail hat keinen Anhang')
        try:
            mailbox.move(msg.uid, 'INBOX/rejected')
        except:
            logging.debug(f'Mail mit dem Betreff {Subject} konnte nicht verschoben werden')
            continue

mailbox.logout()
logging.info(f"{__file__} gestoppt.")