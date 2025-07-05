# calendar
----------

## Goals of the Project:

Create an automated system that:
 - It takes data about applications and deadlines from the Google Spreadsheet.
 - Automatically creates events in Google Calendar.
 - Updates the table by writing the ID of the created events there.
 - Allows the staff of the patent office not to miss deadlines and see everything in one place.
----------

## Components:

 - Google Table with information
 - Python script
 - Google Calendar
---------

## The Working Mechanism

### Step 1: Create a project in Google Cloud Console
Creating a project to access the Google Sheets and Google Calendar API.
Enabling the API.
Getting the credentials file.json (OAuth) or service-account.json.

### Step 2: Setting up the Google Spreadsheet
Creating a table with the necessary columns.
We give access to the service account (or the user, if you use OAuth).
We enter information about applications.

### Step 3: Run the Python Script
When you first run the script, it asks you to log in through the browser → it gives you a link → click "Allow".
After that, the token file is saved.json — it stores the token for subsequent launches.
The script reads data from the table.
Creates an event in the calendar for each row without an Event ID.
Saves the event_id to the table.
-------

## :high_brightness: Results:
We get a program that can be set to a timer so that it automatically transfers all deadlines to a convenient calendar application that clearly shows the deadlines for completing tasks.
