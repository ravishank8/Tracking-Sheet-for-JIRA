Google Sheets - scripts based solutions driven by custom app script to automate both the data collection of Sprint status at the finest grain, and also relevant views for scrum-masters, scrum-of-scrums and engineers.
To code for this tool is present at (https://github.com/ravishank8/Tracking-Sheet-for-JIRA). This location contains the GS code to be imported into any google sheet.
Also a sample google sheet is at: (https://docs.google.com/spreadsheets/d/1G7byXBpfWecynMJ8DMDE7ZTSi9QugiSo9ve1egaVttA/edit?usp=sharing)
The steps are as follows – 
1. The first step is to import the chrome add-on **“JIRA cloud for sheets”** to essentially pull data from JIRA. It supports the following:
* Any JQL needed to pull data from JIRA.
* Ability to select fields to be imported (The default is the columns selected by the user In JIRA)
* Ability to schedule the import – for e.g. once every day, once every hour
We import all this data into a tab by the name “StoriesDump.” This name can be anything you can choose, but this name is used by the code in steps below.

2. The sheet has two key tabs – 
* Calendar: Here you need to mention the working and non-working days of the Sprint
* Metadata: This sheet is to ensure that the code has no dependency on your JIRA field names and JIRA statuses. All this configuration is extracted out of the code. Please review the sheet and appropriately change the field names in JIRA in the column “Data Columns (JIRA)” to reflect the fields that you want to work upon.

3. Now that the basic setup is done, you can choose to rename the generic sheet as per your need. This sheet already has a version of the code.gs in it. The best practice is to open the script editor of your google sheet now, and copy-paste the latest code.gs from github.

4. Now you have already done the following – 

* Got the “JIRA clout for sheets” plugin
* Renamed your file/created a copy
* Imported the latest code from code.gs

5. The sheet is now ready to use. Close the open the sheet once. You will see that the “JIRA Menubar” menu option will re-appear. This is the menubar to execute our custom code.

6. Following are the steps to use the sheet – 
* Import sprint stories into StoriesDump sheet using the “JIRA cloud for sheets.”
* Execute the “Click to Retrieve” option on the JIRA menubar.
* This should populate your StoriesChart tab and the PODStatus tab, which are the most important tabs on the sheet. If you run into any issues with the script, do reach out to the developer on Github 
* If you find some stories missing in the StoriesChart tab, which are there in the StoriesDump tab, the reason for that is very simple - There are only two things that are a must have for stories to be considered for analysis – Story Points (for simplicity we kept 1 SP as 1 person day of work) and Planned End Data. These two are must-have fields to calculate the estimated end date. Any story in JIRA that does not have this gets rejected.

7. StoriesChart Tab has the following important data:
* Story status – In JIRA we have multiple statuses. We map all those statuses to three basic ones – Not started, Complete and In-Progress. Depending on how you manage your QE work, either they are in the same story, or you can create a parallel BDD story for the same. Mapping of these three statuses to the actual JIRA statuses is again maintained in the metadata tab.
*	Delay at each story level – This takes into account each story status – 
  * * Complete (Delay - Completed): Stories can be delayed even if they have already completed.
  * * Not started (Delay in Start - Not yet started): Stories can we delayed even if they have not yet started, based on the effort and planned end date.
  * * In progress stories (Delay In Progress - Due date based): The delay of in progress stories is a combination of delay in start and delay in progress.
Non-working days (including weekends) are taken into account while calculating the delay.

  * * Remaining Effort: This might seem a very obvious thing, but becomes crucial in second week of the Sprint. If all developers are not fully loaded, this data helps give ability to complete the story even though there might have been delays.
  
At this point we have data for each story’s status across all PODs and developers. And this is what forms the data store for the reports.
Now that we have the data, there are 3 views that are the most important. You will see those in the generic sheet too. Once StoriesChart has data, these views will come to life.

8.	The 3 views present 3 different unique pictures – 
 *	**Tracking view** is to answer the following questions. These are typically the first questions to be asked per POD in the scrum of scrums, and in this order too, proceeding to the next one only post clarity on the current question:
  * *  Are you on track? If not, how much are you delayed by?
  * * 	What stories are contributing to the delay? Which developers are delayed?
  * * 	What is the actual delay? (this will be at story level)
Tracking sheet is a PIVOT created from the StoriesChart sheet, answering these exact queries.

![](images/TrackingView.png)

  * **Completion View** – One huge issue during Sprints is that the burn rate is good, but stories do not complete, and most of them get into the build in the second week of a two week Sprint, leading to compromised DOD and delayed closures. Completion view essentially answers the following – 
  * * How are we doing on completing of stories? Whilst effort burn-down presents one picture, how are we doing on stories completion  themselves? Essentially, are we closing stories at the expected pace?
       Again, this is a PIVOT derived from the StoriesChart Data.

![](images/CompletionView.png)

  * The third and final view is the consolidated **POD Status view**, something that is more of a top down throughput view needed by both the engineers, scrums of scrums and steering committee. This after a few sprints get combined with a quality view to present the quality + throughput picture. The throughput view is as such – 
  
![](/images/PODStatusView.png)

**These three views – Tracking Report, Completion Report and POD Status are used to drive the daily scrum of scrums.** In a later section, we will show how we added a quality view to it to present a consolidated velocity + quality view.
**This automation was a primary contributor to us being able to manage with 2 scrum masters across 6 PODs.**
