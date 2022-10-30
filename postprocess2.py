import pandas as pd
import os
import docx
import win32com.client
import os.path
from docx.api import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx2pdf import convert

#relative import testing
from . import common

#extra experiment
import time

#universals - to separate to .common.py local import
#includes certreval and postprocess common
email_signature = ('<b><font color="rgb(0,65,92)"> Peter Soosalu | Coach Development Manager | Yachting New Zealand </font></b> <br>'
    '<b><font color="rgb(0,65,92)">M</b></font> <font color="rgb(0,65,92)">(021) 037 2419 </font>| <b><font color="rgb(0,65,92)">E</font></b> peters@yachting.org.nz <br>'
    '<a href="http://www.yachtingnz.org.nz">Yachtingnz.org.nz</a> | <a href="https://www.facebook.com/YachtingNewZealand/">Facebook</a> | <a href="https://www.facebook.com/NZLSailingTeam/">NZL Sailing Team</a>'
    '<br><br> For the latest news and offerings download the Yachting New Zealand app.'
    '<br><a href="https://apps.apple.com/us/app/yachting-nz-app/id1040333130?amp%3Bls=1&amp%3Bamp%3Bmt=8&amp%3Bl=nl"><img src="https://developer.apple.com/news/images/download-on-the-app-store-badge.png" width="108" height="36" alt="app_store"></a> <a href="https://play.google.com/store/apps/details?id=com.app.p1107GC"><img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT_E6ZM5CF_cPm4tzqW6MpFGm2efBY1QL6v6w&usqp=CAU" width="108" height="36" alt="android_store"></a>'
    '<br><br><img src="https://i.ibb.co/w7Lfz35/99798ce7-9981-4308-8295-3b3fa5386314.jpg" alt="ynz_banner">'
    '<br><br> <font size="-2" color="rgb(220,220,220)"> The content of this e-mail is confidential and may contain copyright information. If you are not the intended recipient, please delete the </font><br> <font size="-2" color="rgb(220,220,220)">message and notify the sender immediately. You should scan this message and any attached files for viruses. We accept no liability for </font><br><font size="-2" color="rgb(220,220,220)"> any loss caused either directly or indirectly by a virus arising from the use of this message or any attached file. Thank you. </font>')

excel_sheet = pd.read_excel("c:/Users/Peters/Downloads/process_these.xlsx")
registrations_expiry = '30 June 2026'

#email attachment variables
safetyRequirementDirectory = ('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\YACHTING NEW ZEALAND SAFETY REQUIREMENT.pdf')
accessingEmbarkDirectory = ('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')

template_path = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Templates\\')
output_path = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs')


def certificate(course_participant, course_type, template_path, output_path):
    if course_type == 'doembark':
        return
    doc = docx.Document(template_path + course_type + '.docx')
    obj_styles = doc.styles
    obj_charstyle = obj_styles.add_style('CommentsStyle', WD_STYLE_TYPE.PARAGRAPH)
    obj_font = obj_charstyle.font
    obj_font.size = Pt(80)
    obj_font.name = 'Freestyle Script'

    for paragraph in doc.paragraphs:
            if 'name_here' in paragraph.text:
                    paragraph.style = doc.styles['CommentsStyle']
                    paragraph.text = course_participant

    doc.save((output_path + '\Yachting New Zealand - ') + course_participant + ('.docx'))
    convert((output_path + '\Yachting New Zealand - ') + course_participant + ('.docx'))
    os.remove((output_path + '\Yachting New Zealand - ') + course_participant + ('.docx'))

    time.sleep(1)

def email(course_participant, course_certificate, participant_email, course_type):
    if course_type == 'doembark':
        body = (("Dear ") + course_participant + (",<br><br>The Yachting New Zealand Coach Course you were on has finished, but have yet to complete the EMBARK online learning portion of the course. The section of the course that you still need to complete is either part of, or all of, the <b>'Coach Yachting 101'</b> section. There are 4 modules: <br> "
                " <b>Embark Introduction </b><br><b>Coach Code of Conduct </b><br><b>Safety Requirements </b><br><b> Quality Coaching </b><br><br>"
                "If you are having trouble accessing EMBARK, I have attached a form that explains how to access the EMBARK learning. <br><br>"
                "Once you have finished, please be sure to email me that you have completed, and I can update the Yachting New Zealand records and get your certificate to you. <br><br>"
                "Alternatively, if you are having trouble accessing EMBARK, please reach out after having tried to login.<br><br>"
                "Regards,<br><br>") + email_signature)
    elif course_type == 'Learn to Sail' or 'Learn to Sail Assistant' or 'Learn to Sail Buddy' or 'Learn to Sail Head':
        if course_type == 'Learn to Sail':
            courseText = "As a Learn to Sail coach you are qualified to teach all aspects of the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Head":
            courseText = "As a Learn to Sail Head coach you are qualified to teach all aspects of the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Assistant":
            courseText = "As an assistant Learn to Sail coach you are qualified to assist with the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Buddy":
            courseText = "As an Buddy Learn to Sail coach you are qualified to assist with the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == 'Keelboat 1' or 'Keelboat 2' or 'Keelboat 3':
            courseText = "As a keelboat coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br> Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 7:1.<br><br>"
        body = (("Dear ") + course_participant + (",<br><br>You have successfully completed the Yachting New Zealand ") + course_type +(" Coach Course. I am pleased to advise that you are now an officially recognised <b> ") + course_type + (" Coach. </b>"
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>")
                 + courseText + ("Your qualification is valid until ") + registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b>") + course_type + ("</b> certificate. You can upgrade to a Learn to sail coach when you turn 18 or working with the feedback the coach developer has given to upgrade your certificate. When you are ready to upgrade, you can by filling out a revalidation from, available to download off the Yachting New Zealand website.   <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "If you are interested in furthering your coaching experience you may be keen to look at some becoming a Race Coach – information can be found under the <b>Coaches</b> page on the Yachting New Zealand website. Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + email_signature)
    elif course_type == 'Race':
        body = (("Dear ") + course_participant + (",<br><br>I am pleased to advise that you are now an officially recognised <b> Race Coach. </b>"
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>"
                "The Race Coach is the first step along the race coach pathway, followed by the Regatta coach, Performance coach, and finally Olympic coach. As a race coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br>"
                "Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 6:1.<br><br>"
                "Your qualification is valid until ") + registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b> Race Coach Certificate.</b> <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + email_signature)
    
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = participant_email
    mail.Subject = 'Yachting New Zealand - Coaching Course Certificate - ' + course_participant 
    mail.HTMLBody = body
    course_certificate = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Yachting New Zealand - ') + course_participant + ('.pdf')
    if course_type == 'doembark':
        mail.attachments.Add(accessingEmbarkDirectory)
    else:
        mail.attachments.Add(course_certificate)
        mail.attachments.Add(safetyRequirementDirectory)

    mail.Send()
    print("Email sent to: " + course_participant)


def revalEmail(course_participant, course_certificate, participant_email, course_type):
    if course_type == 'Learn to Sail' or 'Learn to Sail Assistant' or 'Learn to Sail Buddy' or 'Learn to Sail Head':
        if course_type == 'Learn to Sail':
            courseText = "As a Learn to Sail coach you are qualified to teach all aspects of the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Head":
            courseText = "As a Learn to Sail Head coach you are qualified to teach all aspects of the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Assistant":
            courseText = "As an assistant Learn to Sail coach you are qualified to assist with the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == "Learn to Sail Buddy":
            courseText = "As an Buddy Learn to Sail coach you are qualified to assist with the Yachting New Zealand Learn to Sail Dinghy (Level I and II) program. These levels can be taught at yacht clubs and other Yachting New Zealand affiliated organisations subject to the Yachting New Zealand safety requirements. <br><br>"
        elif course_type == 'Keelboat 1' or 'Keelboat 2' or 'Keelboat 3':
            courseText = "As a keelboat coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br> Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 7:1.<br><br>"
        body = (("Dear ") + course_participant + (",<br><br>After reviewing your revalidation request, I am happy to revalidate your Yachting New Zealand ") + course_type +("certificate. You are now an officially recognised <b> ") + course_type + (" Coach. </b>"
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>")
                 + courseText + ("Your qualification is valid until ") + registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b>") + course_type + ("</b> certificate.<br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "If you are interested in furthering your coaching experience you may be keen to look at some becoming a Race Coach – information can be found under the <b>Coaches</b> page on the Yachting New Zealand website. Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + email_signature)
    elif course_type == 'Race':
        body = (("Dear ") + course_participant + (",<br><br>After reviewing your revalidation request, I am happy to revalidate your Yachting New Zealand <b> Race Coach. </b> certificate."
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>"
                "The Race Coach is the first step along the race coach pathway, followed by the Regatta coach, Performance coach, and finally Olympic coach. As a race coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br>"
                "Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 6:1.<br><br>"
                "Your qualification is valid until ") + registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b> Race Coach Certificate.</b> <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + email_signature)
    
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = participant_email
    mail.Subject = 'Yachting New Zealand - Coaching Course Certificate - ' + course_participant 
    mail.HTMLBody = body
    course_certificate = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Yachting New Zealand - ') + course_participant + ('.pdf')
    if course_type == 'doembark':
        mail.attachments.Add(accessingEmbarkDirectory)
    else:
        mail.attachments.Add(course_certificate)
        mail.attachments.Add(safetyRequirementDirectory)

    mail.Send()
    print("Email sent to: " + course_participant)

#----------------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------------------

#area specific to postprocess
i = 0
while i < len(excel_sheet):
    registrations_name1 = excel_sheet.iloc[i, 3]
    registrations_name = str(registrations_name1)
    registrations_qual1 = excel_sheet.iloc[i, 7]
    registrations_qual = str(registrations_qual1)
    registrations_email1 = excel_sheet.iloc[i, 11]
    registrations_email = str(registrations_email1)

    if registrations_qual == 'Assistant LTS Coach (Dinghy)':
            course_type = 'Learn to Sail Assistant'
    elif registrations_qual == 'LTS Coach (Dinghy)':
            course_type = 'Learn to Sail'
    elif registrations_qual == 'LTS Buddy':
            course_type = 'Learn to Sail Buddy'
    elif registrations_qual == 'Master LTS Coach (Dinghy)':
            course_type = 'Learn to Sail Head'
    elif registrations_qual == 'LTS Coach (Keelboat) Level 1':
            course_type = 'Keelboat 1'
    elif registrations_qual == 'LTS Coach (Keelboat) Level 2':
            course_type = 'Keelboat 2'
    elif registrations_qual == 'LTS Coach (Keelboat) Level 3':
            course_type = 'Keelboat 3'
    elif registrations_qual == 'Race Coach':
            course_type = 'Race'
    else:
            course_type = 'doembark'
    
    # template_path = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Templates\\')
    # output_path = 'c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs'

    certificate(registrations_name, course_type, template_path, output_path)
    email(registrations_name, course_type, registrations_email, course_type)
    i +=1

print("\n DAYUMN, are we already done?! \n That happened WAY too fast. Maybe you should take the rest of the day off. \n")



# #----------------------------------------------------------------------------------------------------------------------------------------------------
# #----------------------------------------------------------------------------------------------------------------------------------------------------
# #CERT REVALIDATION

# course_participant = input("Participant name:")
# print("")
# participant_email = input("Email:")
# print("\nCert type from list below:")
# print("'assistcert' \n'ltscert' \n'headcert' \n'racecert' \n'keel1cert' \n'keel2cert' \n'keel2cert' \n'keel3cert' \n")
# certification_type = input("Cert type: ")

# #will likely have to modify this slightly to work
# certificate(course_participant, certification_type, template_path, output_path)
# revalEmail(course_participant, certification_type, participant_email, course_type)
# print("")
# print(('Revalidation email sent to: ') + course_participant)

# #----------------------------------------------------------------------------------------------------------------------------------------------------
# #----------------------------------------------------------------------------------------------------------------------------------------------------
# #LTS precourse

# print("PRE COURSE EMAIL SYSTEM.")
# print("Make sure the excel sheet is in the downloads folder and renamed 'process_these' \nDo NOT delete the extra columns in the excel document.")
# print("If you have, quit this and just simply blank them out - it will not work properly. \n")
# print("The type of course is either: \n'online,' or \n'traditional' \n")
# print("If you don't enter one of these the program won't work. \n")
# course_type = input("Enter the type of course: ")
# coach_developer_email = input("Coach developer email: ")
# start_date = input("Start date: ")
# personal_cd_touch = input("What do you want to add for the coach developer: ")

# i = 0
# while i < len(excel_sheet):
#         registrations_name1 = excel_sheet.iloc[i, 3]
#         registrations_name = str(registrations_name1)
#         registrations_email1 = excel_sheet.iloc[i, 11]
#         registrations_email = str(registrations_email1)
#         participant_details = [registrations_name, registrations_email]

# #Generic course related details trying to loop
#         course_participant = participant_details[0]
#         participant_email = participant_details[1]

#         #outlook mail portion
#         outlook = win32com.client.Dispatch('outlook.application')
#         mail = outlook.CreateItem(0)

# #Online version
#         if course_type == 'online':
#                 mail.To = participant_email
#                 mail.Subject = "Yachting New Zealand - Coaching Course Details"
#                 mail.HTMLBody = ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
#                         "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
#                         "- Online learning sections (BRACKEN)<br><br>- Interactive zoom sessions on coaching (3 x 2 hours each)<br><br>- On-water section to demonstrate safe coaching and sailing abilities<br><br><br>"
#                         "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! <br><br>"
#                         "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
#                         "<b><u>Zoom sessions</b></u><br><br> For the online sessions, Yachting New Zealand uses zoom. If you are new to zoom, make an account and try logging on to the zoom invite BEFORE the course dates. As a coach, you are responsible for getting yourself to the course on time in person, and this is no different. <br><br> The zoom sessions will be emailed you to separately. The first zoom session is ") + start_date + ("<br><br>It is your responsibility to plan and make sure you are ready and available for all 3 sessions. <u>You must attend all 3 sessions to get your certification.</u> Control the controllable, and plan ahead!<br><br>"
#                         "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
#                         "<b><u>On the water component</u></b><br><br> Once you have finished the online sections, you will either have to follow up with your coach developer, or with a Regional Development Manager for the on water portion. <br> If you are planning to bring your own boat to the course, please let the coach developer know. If you are planning to use a Learn to Sailboat from the club, make sure you have organized this beforehand with the club.<br><br><br>"
#                         "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
#                         "Any questions specifically on how the course will run for the sessions, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
#                         "Regards,<br><br>") + email_signature
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
#                 mail.Send()
#                 print(('Precourse email to: ') + course_participant)
#                 i +=1

# #Traditional version
#         elif course_type == 'traditional':
#                 mail.To = participant_email
#                 mail.Subject = "Yachting New Zealand - Coaching Course Details"
#                 mail.HTMLBody = ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
#                         "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
#                         "-	Online learning sections (BRACKEN)<br><br>-	On-water section to demonstrate safe coaching and sailing abilities<br><br>- On-shore theory and overview of the LTS 1 & 2 courses<br><br><br>"
#                         "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! You will need to complete the <b>Coach Yachting 101</b> section, while the Learn to Sail manual is a useful resource that you will use during the Learn to Sail course. <br><br>"
#                         "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
#                         "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
#                         "<b><u>Boats on the water</u></b><br><br> If you are planning to bring your own boat to the course to demonstrate your sailing ability, please let the coach developer know at the email below.<br><br><br>"
#                         "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
#                         "Any questions specifically on how the course will run for the day, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
#                         "Regards,<br><br>") + email_signature
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
#                 mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
#                 mail.Send()

#                 print(('Precourse email to: ') + course_participant)
#                 i +=1

# #Email for coach developer
# mail = outlook.CreateItem(0)
# mail.To = coach_developer_email
# mail.Subject = "Yachting New Zealand - Coaching Course Details"
# mail.HTMLBody = personal_cd_touch + ("<br><br><br><br><br>") + ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
#         "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
#         "-	Online learning sections (BRACKEN)<br><br>-	On-water section to demonstrate safe coaching and sailing abilities<br><br>- On-shore theory and overview of the LTS 1 & 2 courses<br><br><br>"
#         "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! <u>You will need to complete the Coach Yachting 101 section </u>to finish the course, while the Learn to Sail manual is a useful resource that you will use during the Learn to Sail course. <br><br>"
#         "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
#         "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
#         "<b><u>Boats on the water</u></b><br><br> If you are planning to bring your own boat to the course to demonstrate your sailing ability, please let the coach developer know at the email below. If you are planning to use a Learn to Sail boat, make sure you have organized this beforehand with the club contact.<br><br><br>"
#         "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
#         "Any questions specifically on how the course will run for the day, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
#         "Regards,<br><br>") + email_signature
# mail.attachments.Add('c:\\Users\Peters\Downloads\process_these.xlsx')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
# mail.Send()

# print(('CD email to: ') + coach_developer_email)


# #---------------------------------------------------------------------------------------------------------------------------------------------------
# #---------------------------------------------------------------------------------------------------------------------------------------------------
# #single cert script
# course_participant = input("Participant name:")
# participant_email = input("\nEmail:")
# print("\nCert type from list below:")
# print("'assistcert' \n'ltscert' \n'headcert' \n'racecert' \n'keel1cert' \n'keel2cert' \n'keel2cert' \n'keel3cert' \n")
# #or this one?
# print("Cert type from list below: \nBuddy = b \nAssistant = a \nLTS =  l \nHead = h \nRace = ra \nKeelboat 1 = k1 \nKeelboat 2 = k2 \nKeelboat 3 = k3")
# certification_type = input("Cert type: ")

# certificate(course_participant, certification_type, template_path, output_path)
# email(course_participant, certification_type, participant_email, 'courseTypePlaceholder')