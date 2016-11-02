# coding=utf-8

import imaplib
import email
from email import generator
import os
import re
import zipfile
import argparse

def get_mailaccount(username, password, server, folder):
    """Logs into an E-mailaccount using imap and downloads all mails and tghe folderstructure as well!"""
    conn = imaplib.IMAP4_SSL(server)
    conn.login(username, password)

    # Create a folder to put the whole stuff in
    os.mkdir(folder)

    # Creates an iterator with all mailboxes
    mailbox_folders = (f.decode().split(' "." ')[-1] for f in conn.list()[1])
    
    for mailbox in mailbox_folders:        
        # Creates The current Mailbox
        current_folder = os.path.join(folder, *mailbox.strip("\"\'").split("."))
        os.mkdir(current_folder)

        # Selects and reads the mailbox
        conn.select(mailbox=mailbox, readonly=True)

        typ, data = conn.search(None, 'ALL')

        # Saving the result to the current folder
        for mail in data[0].split():
            typ, data = conn.fetch(mail, '(BODY.PEEK[])')

            # Creates the text of the mail to extract sender and subject
            text = data[0][1].decode()
            mailtext = email.message_from_string(text)
            if "subject" in mailtext:
                regex = re.compile('[^a-zA-Z0-9,äöüÄÖÜ ]')
                subject = regex.sub('', mailtext["subject"])
            else:
                subject = "unknown"                

            # Renames for avoiding of doublenames
            savename = subject + ".eml"
            i = 1
            while os.path.isfile(savename):
                savename = filename + "." + str(i) + ".eml"
                i += 1
            filename = "{}\\{}".format(current_folder, savename)

            # Checks if the mail contains an attachement and downloads it then
            if mailtext.get_content_maintype() == 'multipart':
                for part in mailtext.walk():
                    if part.get_content_maintype() == 'multipart':
                            continue
                    if part.get('Content-Disposition') is None:
                            continue

                    # Creates a folder for the attachments
                    attachmentpath = os.path.join(current_folder, "attachmentfolder- " + savename)
                    if not os.path.exists(attachmentpath):
                        os.mkdir(attachmentpath)

                    # Two Attachments with the same name? No props!
                    attachmentname = part.get_filename()
                    i = 1
                    while os.path.isfile(os.path.join(attachmentpath, attachmentname)):
                        attachmentname = str(i) + "-" + attachmentname
                        i += 1
                    savepath = os.path.join(attachmentpath, attachmentname)

                    # Save attachment
                    with open(savepath, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                        fp.close()

            # Saves the result
            with open(filename, "w") as target:
                gen = generator.Generator(target)
                gen.flatten(mailtext)
            
    conn.close()
    conn.logout()


def zip_dir(inputDir, outputZip):
    '''Zip up a directory and preserve symlinks and empty directories'''
    # Credits go to kgn/ZipDir.py
    zipOut = zipfile.ZipFile(outputZip, 'w', compression=zipfile.ZIP_DEFLATED)
    rootLen = len(os.path.dirname(inputDir))

    # For recursion reasons
    def archiveDirectory(parentDirectory):
        contents = os.listdir(parentDirectory)

        # If it is empty - make it anyway
        if not contents:
            archiveRoot = parentDirectory[rootLen:].replace('\\', '/').lstrip('/')
            zipInfo = zipfile.ZipInfo(archiveRoot+'/')
            zipOut.writestr(zipInfo, '')

        # All objects need to be written in the zipfile
        for item in contents:
            fullPath = os.path.join(parentDirectory, item)

            # If it is a folder, call archiveDirectory recursiv
            if os.path.isdir(fullPath) and not os.path.islink(fullPath):
                archiveDirectory(fullPath)
            else:
                archiveRoot = fullPath[rootLen:].replace('\\', '/').lstrip('/')
                zipOut.write(fullPath, archiveRoot, zipfile.ZIP_DEFLATED)
                
    archiveDirectory(inputDir)
    zipOut.close()


    
if __name__ == "__main__":
    
    parser = argparse.ArgumentParser(description='Downloads all mails using IMAP')
    parser.add_argument('server', metavar='S', type=str,
                        help='enter the IMAP - server you want to pull the mails from.')
    parser.add_argument('username', metavar='U', type=str, help='your username')
    parser.add_argument('password', metavar='P', type=str, help='your password')
    parser.add_argument('--folder', metavar='F', type=str, default='Mails',
                        help='the place you want to store your results')
    args = parser.parse_args()

    get_mailaccount(args.username, args.password, args.server, args.folder)
    zip_dir(os.path.join(os.getcwd(),args.folder), os.path.join(os.getcwd(),args.folder + ".zip"))
    
