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
 :white_small_square: Creating a project to access the Google Sheets and Google Calendar API.
 :white_small_square: Enabling the API.
 :white_small_square: Getting the credentials file.json (OAuth) or service-account.json.

### Step 2: Setting up the Google Spreadsheet
 :white_small_square: Creating a table with the necessary columns.
 :white_small_square: We give access to the service account (or the user, if you use OAuth).
 :white_small_square: We enter information about applications.


:white_small_square:
-------

## :high_brightness: Results:
We get a program that can be set to a timer so that it automatically transfers all deadlines to a convenient calendar application that clearly shows the deadlines for completing tasks.
