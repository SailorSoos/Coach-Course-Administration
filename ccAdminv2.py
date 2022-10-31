#import os
#import os.path
#import time

#import docx
#from docx.api import Document
#from docx.enum.style import WD_STYLE_TYPE
#from docx.shared import Pt
#from docx2pdf import convert
#import pandas as pd
import win32com.client

from common import certificate, email, revalEmail, Universal

print("Welcome to the new and improved coaching automation system! \n You have a few options of what you can do. Type an option from the list below.")
print("As always, the course list of participants must be in your downloads folder, and renamed 'process_these.xlsx'")
print("The options are: \n Post process (pp) \n Revalidation (r) \n Pre Course (pc) \n Single certification (sc)")
selected_option = input("What do you want to do?! Please tell me :): ")
print("FINNNNNEEEEE! I'm doing it as quick as I can. Prepare to be yelled at if you haven't selected one of the above options. \n \n I'll even update you on how we're doing, bud.")

#----------------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------------------
#POST PROCESS
if selected_option == 'pp':
    i = 0
    while i < len(Universal.excel_sheet):
        registrations_name1 = Universal.excel_sheet.iloc[i, 3]
        registrations_name = str(registrations_name1)
        registrations_qual1 = Universal.excel_sheet.iloc[i, 7]
        registrations_qual = str(registrations_qual1)
        registrations_email1 = Universal.excel_sheet.iloc[i, 11]
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

        certificate(registrations_name, course_type, Universal.template_path, Universal.output_path)
        email(registrations_name, course_type, registrations_email, course_type)
        i +=1

    print("\n DAYUMN, are we already done?! \n That happened WAY too fast. Maybe you should take the rest of the day off. \n")


#----------------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------------------
#CERT REVALIDATION
elif selected_option == 'r':
    course_participant = input("Participant name:")
    print("")
    participant_email = input("Email:")
    print("\nCert type from list below:")
    print("'assistcert' \n'ltscert' \n'headcert' \n'racecert' \n'keel1cert' \n'keel2cert' \n'keel2cert' \n'keel3cert' \n")
    certification_type = input("Cert type: ")

    certificate(course_participant, certification_type, Universal.template_path, Universal.output_path)
    revalEmail(course_participant, certification_type, participant_email, certification_type)
    print("")
    print(('Revalidation email sent to: ') + course_participant)

#----------------------------------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------------------------------
#LTS precourse
elif selected_option == 'pc':
    print("The type of course is either: \n'online,' or \n'traditional' \n")
    course_type = input("Enter the type of course: ")
    coach_developer_email = input("Coach developer email: ")
    start_date = input("Start date: ")
    personal_cd_touch = input("What do you want to add for the coach developer: ")

    i = 0
    while i < len(Universal.excel_sheet):
            registrations_name1 = Universal.excel_sheet.iloc[i, 3]
            registrations_name = str(registrations_name1)
            registrations_email1 = Universal.excel_sheet.iloc[i, 11]
            registrations_email = str(registrations_email1)
            participant_details = [registrations_name, registrations_email]

    #Generic course related details trying to loop
            course_participant = participant_details[0]
            participant_email = participant_details[1]

            #outlook mail portion
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            if course_type == 'online':
                    mail.To = participant_email
                    mail.Subject = "Yachting New Zealand - Coaching Course Details"
                    mail.HTMLBody = ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
                            "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
                            "- Online learning sections (BRACKEN)<br><br>- Interactive zoom sessions on coaching (3 x 2 hours each)<br><br>- On-water section to demonstrate safe coaching and sailing abilities<br><br><br>"
                            "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! <br><br>"
                            "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
                            "<b><u>Zoom sessions</b></u><br><br> For the online sessions, Yachting New Zealand uses zoom. If you are new to zoom, make an account and try logging on to the zoom invite BEFORE the course dates. As a coach, you are responsible for getting yourself to the course on time in person, and this is no different. <br><br> The zoom sessions will be emailed you to separately. The first zoom session is ") + start_date + ("<br><br>It is your responsibility to plan and make sure you are ready and available for all 3 sessions. <u>You must attend all 3 sessions to get your certification.</u> Control the controllable, and plan ahead!<br><br>"
                            "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
                            "<b><u>On the water component</u></b><br><br> Once you have finished the online sections, you will either have to follow up with your coach developer, or with a Regional Development Manager for the on water portion. <br> If you are planning to bring your own boat to the course, please let the coach developer know. If you are planning to use a Learn to Sailboat from the club, make sure you have organized this beforehand with the club.<br><br><br>"
                            "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
                            "Any questions specifically on how the course will run for the sessions, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
                            "Regards,<br><br>") + Universal.email_signature
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
                    mail.Send()
                    print(('Precourse email to: ') + course_participant)
                    i +=1

            else:
                    mail.To = participant_email
                    mail.Subject = "Yachting New Zealand - Coaching Course Details"
                    mail.HTMLBody = ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
                            "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
                            "-	Online learning sections (BRACKEN)<br><br>-	On-water section to demonstrate safe coaching and sailing abilities<br><br>- On-shore theory and overview of the LTS 1 & 2 courses<br><br><br>"
                            "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! You will need to complete the <b>Coach Yachting 101</b> section, while the Learn to Sail manual is a useful resource that you will use during the Learn to Sail course. <br><br>"
                            "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
                            "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
                            "<b><u>Boats on the water</u></b><br><br> If you are planning to bring your own boat to the course to demonstrate your sailing ability, please let the coach developer know at the email below.<br><br><br>"
                            "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
                            "Any questions specifically on how the course will run for the day, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
                            "Regards,<br><br>") + Universal.email_signature
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
                    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
                    mail.Send()

                    print(('Precourse email to: ') + course_participant)
                    i +=1

    #Email for coach developer
    mail = outlook.CreateItem(0)
    mail.To = coach_developer_email
    mail.Subject = "Yachting New Zealand - Coaching Course Details"
    mail.HTMLBody = personal_cd_touch + ("<br><br><br><br><br>") + ("Good day soon to be coaches,<br><br>Hope you are all looking forward to the coaching course starting on <u>") + start_date + (".</u> I am reaching out to give you a brief overview of the course, as well as a reminder about key dates and what needs to be covered to complete the course. "
            "As we have moved online for a portion of the course, the new Learn to Sail course has 3 segments.<br><br>"
            "-	Online learning sections (BRACKEN)<br><br>-	On-water section to demonstrate safe coaching and sailing abilities<br><br>- On-shore theory and overview of the LTS 1 & 2 courses<br><br><br>"
            "<u><b>Online learning access</u></b><br><br>The way you would have registered for the course is how you will access the <u>online learning</u>, as well as the <u>online Learn to Sail manual.</u><br><br>The manual is a massive resource for you as coaches! <u>You will need to complete the Coach Yachting 101 section </u>to finish the course, while the Learn to Sail manual is a useful resource that you will use during the Learn to Sail course. <br><br>"
            "Attached is a “how to login to BRACKEN” troubleshooting document if you are stumbling on how to access.<br><br> If you are interested in having a hard copy of the manual, Embark online learning does offer the option to print slides for later use, which is available on the top right corner when you log in to Embark. For a simple “how to print” section, you can look through the attachments on this email. <br><br><br>"
            "I have attached a few coaching resources that are useful to print and have ready for the course as well.<br><br><br>"
            "<b><u>Boats on the water</u></b><br><br> If you are planning to bring your own boat to the course to demonstrate your sailing ability, please let the coach developer know at the email below. If you are planning to use a Learn to Sail boat, make sure you have organized this beforehand with the club contact.<br><br><br>"
            "<b><u>Other requirements</u></b><br><br>Be sure to bring the following to the course<br><br> - Any gear necessary to keep warm on the water as a coach<br><br> - Lifejacket<br><br> - Whistle<br><br> - Wet notes and pencil<br><br>"
            "Any questions specifically on how the course will run for the day, make sure to get in contact with ") + coach_developer_email + ("<br><br>"
            "Regards,<br><br>") + Universal.email_signature
    mail.attachments.Add('c:\\Users\Peters\Downloads\process_these.xlsx')
    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
    mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
    mail.Send()

    print(('CD email to: ') + coach_developer_email)


#---------------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------------------------------------------------------------------
#single cert script
elif selected_option == 'pc':
    course_participant = input("Participant name:")
    participant_email = input("\nEmail:")
    print("\nCert type from list below:")
    print("'assistcert' \n'ltscert' \n'headcert' \n'racecert' \n'keel1cert' \n'keel2cert' \n'keel2cert' \n'keel3cert' \n")
    #or this one?
    print("Cert type from list below: \nBuddy = b \nAssistant = a \nLTS =  l \nHead = h \nRace = ra \nKeelboat 1 = k1 \nKeelboat 2 = k2 \nKeelboat 3 = k3")
    certification_type = input("Cert type: ")

    certificate(course_participant, certification_type, Universal.template_path, Universal.output_path)
    email(course_participant, certification_type, participant_email, 'courseTypePlaceholder')

else:
    print("LOOK BUD, I AM YELLING AT YOU BECAUSE ONE OF THE ABOVE FUNCTIONS WAS NOT SELECTED. Close me down and let's try again :)")