#!/usr/bin/env python3
#-*-coding:utf8;-*-
############################################################################
#   MailCatcher by Sven Gaechter Allround-Service.biz                      #
#                                                                          #
#   History:                                                               #
#   0.1: Initial quick and dirty Code                                      #
#   1.0: First Release                                                     #
#                                                                          #
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
moveMail = config['General']['moveMailTo']

##### Log-Handler #####
logging.basicConfig(filename=os.path.join(os.path.dirname(__file__),'Mailcatcher.log'),format='%(asctime)s - %(message)s', level=config['General']['loglevel'])

def processMailbox():
    with MailBox(mailServer).login(login, password, initial_folder='INBOX') as mailbox:
        logging.debug(sg.subject for msg in mailbox.fetch())

        for msg in mailbox.fetch(Q(from_=senderAddr)):
            logging.debug(f'Betreff: {msg.subject} von {msg.date}')
            if len(msg.attachments) > 0:
                attachment = [i[0] for i in msg.attachments] + [i[1] for i in msg.attachments]

                if fileExt in attachment[0]:
                    logging.info(f'{attachment[0]} wird heruntergeladen...')
                    f = open(os.path.join(attPath, attachment[0]), "wb") #destFilename If the Filename will not be replaced
                    f.write(attachment[1])
                    f.close()
                    try:
                        mailbox.move(msg.uid, moveMail)
                        #mailbox.delete(msg.id)
                        logging.debug(f'Mail verschoben.')
                    except:
                        logging.debug(f'Mail mit dem Betreff {msg.subject} konnte nicht verschoben werden')
                        continue
                else:
                    logging.info(f'{attachment[0]} entspricht nicht dem Suchbegriff {fileExt}. Datei wurde nicht heruntergeladen')
                    try:
                        mailbox.move(msg.uid, 'INBOX/rejected')
                        logging.debug(f'Mail verschoben.')
                        continue
                    except:
                        logging.debug(f'Mail mit dem Betreff {msg.subject} konnte nicht verschoben werden')
                        continue
            else:
                logging.debug(f'Mail hat keinen Anhang')
                try:
                    mailbox.move(msg.uid, 'INBOX/rejected')
                except:
                    logging.debug(f'Mail mit dem Betreff {msg.subject} konnte nicht verschoben werden')
                    continue

if __name__ == "__main__":
    logging.info('-------------------------')
    logging.info(f"{__file__} gestartet...")
    processMailbox()
    logging.info(f"{__file__} beendet.")