# Setup
## Install dependencies
run the command "pip install -r requirements.txt".
## Get OAuth2 token
To authenticate your requests to the canvas API, you'll need an OAuth2 access token. Open canvas and select the "account" button on the left side. Then select Profile, and then click settings on the left. Scroll down until you see "approved integrations". Click New access token, and set the purpose and expiry date. Finally, copy the token into a text file and save it. You will need to set the global AUTH_PATH variable in main.py to the path to this text file.
# Usage
To choose which lab to grade, you will need to set the global ASSIGNMENT variable in main.py to the id of the lab you want. You can find this by opening the lab in speedgrader and looking at the number after assignments in the URL bar.
## Input format
<p>The program expects an <b>xlsx</b> file with the following format. The students' names should be in the second column in middle last, first format. The middle name is optional. The tops of the first two columns are ignored. From the third column onwards, the first row should contain the week of the lab (for example, the lab for the week of 9/11 should have a header like "week 9/11"). These columns represent whether or not the student was present that week's lab. The value will be interpreted as present if the column contains the string "present" or "yes", and absent if it is "absent" or "no". The first column is ignored (I used the lists for the labs Dr. Yarrington sent, which had their emails in that column).</p>
Save this in an <b>xlsx</b> sheet. Set the global ATTENDANCE_PATH variable in main.py to the path to the sheet, and ATTENDANCE_SHEET to the name of the sheet in the workbook containing the presence/absence of students.