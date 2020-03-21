# Corona Lunch Partner

Made a small google script that sends an email to people that want to eat lunch together online during the COVID-19 lockdown.

## Howto

### 1. Create a new Google Sheets file with a script
Name the file "Lunch"

Hit Tools -> Script Editor -> Copy/Paste lunch.gs

### 2. Create 2 google forms
Create 2 google forms, 1 to sign up and 1 to cancel.
Make sure to tick of the box that says "Collect email addresses" under Settings in both forms.
Make sure that the responses from both forms are stored in Lunch file by using the "Select response destination"

#### 2.1 Response form
Question 1: Name
Question 2: Write a little bit about yourself

Name the sheet "responses"

#### 2.2 Cancelleation form
Question 1: Name

Name the sheet "cancellations"

### 3. Set the trigger
See [triggerDaily() at the bottom of lunch.gs](lunch.gs)

### 4. Run the trigger function
Everything should now be functioning.
Make sure that `MailApp.sendEmail(email, subject, message);` is uncommented
