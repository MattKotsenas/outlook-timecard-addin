Outlook Timecard Add-In
=======================

A small Outlook Add-In to help with tracking your hours / timecard.

The Idea
--------

Let's assume you usually start and end your day sending a few emails. Let's also assume that you'd like to know when your day started and ended, perhaps to fill out your time card.

How It Works
------------

1. Open the add-in via the Add-In tab in Outlook's ribbon
2. Select the start and end dates in the calendar control
3. Hit calculate (this may take a few seconds depending on the amount of mail you have)

The add-in looks at your sent mail and makes a list of the earliest and latest mail for each day in the range. From there you know roughly when your day started and ended!

Prerequisites
-------------

* Outlook 2010 or 2013
* Windows 7 or Windows 8
* .NET 4 (ClickOnce installs this if needed)

Deploying
---------

1. Build the project in Visual Studio
2. Open the Project Properties
3. Go to the Publish tab
4. Click 'Publish Now'
5. Run the ClickOnce installer