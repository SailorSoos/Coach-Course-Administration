#Put the excel export to the path here, change 'process_these.xlsx' to the filename, or rename the filename in the download folder to 'process_these'
import pandas as pd
import os
import re
import win32com.client
import os.path
excel_sheet = pd.read_excel("c:/Users/Peters/Downloads/process_these.xlsx")

print("PRE COURSE EMAIL SYSTEM. ")
print("Make sure the excel sheet is in the downloads folder and renamed 'process_these' \nDo NOT delete the extra columns in the excel document.")
print("If you have, quit this and just simply blank them out - it will not work properly. \n")
print("The type of course is either:")
print("'online,' or \n'traditional' \n")
print("If you don't enter one of these the program won't work. \n")
course_type = input("Enter the type of course: ")
coach_developer_email = input("Coach developer email: ")
start_date = input("Start date: ")
personal_cd_touch = input("What do you want to add for the coach developer: ")

i = 0
while i < len(excel_sheet):
        registrations_name1 = excel_sheet.iloc[i, 3]
        registrations_name = str(registrations_name1)
        registrations_email1 = excel_sheet.iloc[i, 11]
        registrations_email = str(registrations_email1)
        participant_details = [registrations_name, registrations_email]

#Generic course related details trying to loop
        course_participant = participant_details[0]
        participant_email = participant_details[1]

#Email signature details
        line_1 = ('<b><font color="rgb(0,65,92)"> Peter Soosalu | Coach Development Manager | Yachting New Zealand </font></b> <br>')
        line_2 = ('<b><font color="rgb(0,65,92)">M</b></font> <font color="rgb(0,65,92)">(021) 037 2419 </font>| <b><font color="rgb(0,65,92)">E</font></b> peters@yachting.org.nz <br>')
        line_3 = ('<a href="http://www.yachtingnz.org.nz">Yachtingnz.org.nz</a> | <a href="https://www.facebook.com/YachtingNewZealand/">Facebook</a> | <a href="https://www.facebook.com/NZLSailingTeam/">NZL Sailing Team</a>') 
        line_4 = ('<br><br> For the latest news and offerings download the Yachting New Zealand app.') 
        line_5 = ('<br><a href="https://apps.apple.com/us/app/yachting-nz-app/id1040333130?amp%3Bls=1&amp%3Bamp%3Bmt=8&amp%3Bl=nl"><img src="https://developer.apple.com/news/images/download-on-the-app-store-badge.png" width="108" height="36" alt="app_store"></a> <a href="https://play.google.com/store/apps/details?id=com.app.p1107GC"><img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT_E6ZM5CF_cPm4tzqW6MpFGm2efBY1QL6v6w&usqp=CAU" width="108" height="36" alt="android_store"></a>')
        line_6 = ('<br><br><img src="https://i.ibb.co/w7Lfz35/99798ce7-9981-4308-8295-3b3fa5386314.jpg" alt="ynz_banner">')
        line_7 = ('<br><br> <font size="-2" color="rgb(220,220,220)"> The content of this e-mail is confidential and may contain copyright information. If you are not the intended recipient, please delete the </font><br> <font size="-2" color="rgb(220,220,220)">message and notify the sender immediately. You should scan this message and any attached files for viruses. We accept no liability for </font><br><font size="-2" color="rgb(220,220,220)"> any loss caused either directly or indirectly by a virus arising from the use of this message or any attached file. Thank you. </font>')
        email_signature = (line_1 + line_2 + line_3 + line_4 + line_5 + line_6 + line_7)

        #outlook mail portion
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

#Online version
        if course_type == 'online':

#Create the email 
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
                        "Regards,<br><br>") + email_signature
#Attachments
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
                mail.Send()

                print(('Precourse email to: ') + course_participant)
                i +=1


#Traditional version
        elif course_type == 'traditional':
            
#Create the email 
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
                        "Regards,<br><br>") + email_signature

#Attachments
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
                # mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
                mail.Send()

                print(('Precourse email to: ') + course_participant)
                i +=1
        else:
            print("Whoops")
            i +=1


#Email for coach developer
mail = outlook.CreateItem(0)

#Create the email 
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
        "Regards,<br><br>") + email_signature

#Attachments
mail.attachments.Add('c:\\Users\Peters\Downloads\process_these.xlsx')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Accessing Embark.pdf')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\Guide - how to print in Embark.docx')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Back End to write on.doc')
# mail.attachments.Add('c:\\Users\Peters\Documents\CDM\Certificates\Attachments\LTS Coach Manual - Lesson Plan on and off water.doc')
mail.Send()

print(('CD email to: ') + coach_developer_email)