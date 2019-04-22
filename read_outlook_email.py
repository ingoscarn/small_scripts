import imaplib

email_user = 'user_outlook'
email_pass = 'pass_outlook'
contador = 0

M = imaplib.IMAP4_SSL('outlook.office365.com', 993)
M.login(email_user, email_pass)
M.select()

typ, message_numbers = M.search(None, '(UNSEEN)','(SUBJECT "Texto del subject")') 

for num in message_numbers[0].split():
    typ, data = M.fetch(num,  '(UID BODY[TEXT])')
    contador= contador + 1
    fh=open("/home/dbadmin/read_email/staging/mail.txt."+str(contador),"w")
    fh.write(str(data))
    fh.close()
M.close()
M.logout()
