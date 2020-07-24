# Certificate-Generator
- This program generates e-Certificates and mails them to participants.
- It uses PIL ImageDraw to write the name of student on the certificate.
- It uses MIME to edit the mailing details and add attachment(certificate) to the participants.
- SMTP is used to mail.

## To generate certificates :
- First you need to add the gmail address and password of the sender in the variables fromEmail and pwd respectively.
- Enter the word document's name, where the details of participants are stored, in df_students.
- Now enter the template of the certificate to be used in Image.open function.
- The x,y coordinates on the certificate template from where the name is to be started writing.
- Change the text in msg['Subject] to change the subject of the mail.
- The body variable can be editied as per requirement to change the body of the mail to be sent.
