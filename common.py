import os
import os.path
import time

import docx
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx2pdf import convert
import pandas as pd
import win32com.client


#to be organised and converted over to classes and methods
class Universal():
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
                "Regards,<br><br>") + Universal.email_signature)
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
                 + courseText + ("Your qualification is valid until ") + Universal.registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b>") + course_type + ("</b> certificate. You can upgrade to a Learn to sail coach when you turn 18 or working with the feedback the coach developer has given to upgrade your certificate. When you are ready to upgrade, you can by filling out a revalidation from, available to download off the Yachting New Zealand website.   <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "If you are interested in furthering your coaching experience you may be keen to look at some becoming a Race Coach – information can be found under the <b>Coaches</b> page on the Yachting New Zealand website. Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + Universal.email_signature)
    elif course_type == 'Race':
        body = (("Dear ") + course_participant + (",<br><br>I am pleased to advise that you are now an officially recognised <b> Race Coach. </b>"
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>"
                "The Race Coach is the first step along the race coach pathway, followed by the Regatta coach, Performance coach, and finally Olympic coach. As a race coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br>"
                "Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 6:1.<br><br>"
                "Your qualification is valid until ") + Universal.registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b> Race Coach Certificate.</b> <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + Universal.email_signature)
    
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = participant_email
    mail.Subject = 'Yachting New Zealand - Coaching Course Certificate - ' + course_participant 
    mail.HTMLBody = body
    course_certificate = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Yachting New Zealand - ') + course_participant + ('.pdf')
    if course_type == 'doembark':
        mail.attachments.Add(Universal.accessingEmbarkDirectory)
    else:
        mail.attachments.Add(course_certificate)
        mail.attachments.Add(Universal.safetyRequirementDirectory)

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
                 + courseText + ("Your qualification is valid until ") + Universal.registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b>") + course_type + ("</b> certificate.<br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "If you are interested in furthering your coaching experience you may be keen to look at some becoming a Race Coach – information can be found under the <b>Coaches</b> page on the Yachting New Zealand website. Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + Universal.email_signature)
    elif course_type == 'Race':
        body = (("Dear ") + course_participant + (",<br><br>After reviewing your revalidation request, I am happy to revalidate your Yachting New Zealand <b> Race Coach. </b> certificate."
                "As part of an effort to reduce the carbon footprint certificates make, I have attached your certificate as a pdf for you to download. If you would like a hard copy of the certificate, please let me know, along with your mailing address, and I can make sure it gets to the right place. <br><br>"
                "The Race Coach is the first step along the race coach pathway, followed by the Regatta coach, Performance coach, and finally Olympic coach. As a race coach, mentoring and supporting other coaches in your area is not only encouraged, but can help improve your own coaching experience by sharing ideas and seeing new ones. <br><br>"
                "Your qualification does not have a set required sailor to coach ratio, but it is recommended to have a max ratio of 6:1.<br><br>"
                "Your qualification is valid until ") + Universal.registrations_expiry + (" at which time you will need to revalidate. Please find enclosed your <b> Race Coach Certificate.</b> <br><br>"
                "Although not mandatory, we do highly recommend that the following courses be added to your qualifications: RYA Powerboat level 2, current first aid certificate and you may also consider a VHF Operators Certificate. Information can be found on the Yachting New Zealand and Coastguard Boating Education websites.<br><br>"
                "Also check out the Yachting New Zealand Coaches Forum Facebook group.<br><br>"
                "Congratulations again on attaining this qualification. Please do not hesitate to contact me at any time if you require any additional assistance or guidance during your coaching.<br><br>"
                "Regards,<br><br>") + Universal.email_signature)
    
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = participant_email
    mail.Subject = 'Yachting New Zealand - Coaching Course Certificate - ' + course_participant 
    mail.HTMLBody = body
    course_certificate = ('c:\\Users\Peters\Documents\CDM\Certificates\Auto_created_certs\Yachting New Zealand - ') + course_participant + ('.pdf')
    if course_type == 'doembark':
        mail.attachments.Add(Universal.accessingEmbarkDirectory)
    else:
        mail.attachments.Add(course_certificate)
        mail.attachments.Add(Universal.safetyRequirementDirectory)

    mail.Send()
    print("Email sent to: " + course_participant)