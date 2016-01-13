####FTC 6055 GearTicks
####Caleb Sander

#Introduction
We are FTC Team 6055, the GearTicks, from Lincoln, Massachusetts. This is our fourth year doing FTC, and all of us were part of FLL teams for a few years before. We won the Inspire Award at the state championship in 2014 and 2015 and the Eastern Super-regional in 2015, going on to the World Championship both years. The code and this document were primarily written by Caleb Sander, the lead programmer of the team. If you need to contact the team, send us an e-mail at ftc6055gearticks@gmail.com.
We have been working on this script since the end of September 2015 and began using it on October 9. We are continuing to add small features and make small changes to it, but it seems very stable by now. You can check out this [example document](https://drive.google.com/file/d/0BxTVo0Nz6n-tM0UtNmc1SmFVczQ/view?usp=sharing) that it produced.

#The Problems
For the last two years, we wrote all our Engineering Notebook (EN) entries by hand and manually inserted all images and drawings. Although we were mostly happy with our EN, we found a few issues inherent to the way we created it:
- **Participation**: We almost always passed around the EN at the end of each meeting and asked everyone to write a short reflection about what they had done. Because many people were in a rush or didn't want to spend a long time writing, people often wrote very little and the entries didn't accurately reflect what we had done at the meeting.
- **Reference**: The EN was not in a very usable form. It was a hassle to look back through it, and we often found that we hadn't recorded data that we wished we had.
- **Legibility**: The fact that everything was handwritten and often in a hurry meant that much of what we had written couldn't easily be used by either the team or judges.
- **Images**: Our image insertion process required someone to take a picture and send it to someone with a computer that could print to our printer. We then had to put the images on glossy paper, cut them out, and caption them. This process was time-consuming and required a lot of knowledge of the process that few of us had (see **Formatting**).
- **Formatting**: The procedure for setting up a page to used for a meeting's entry required knowing all the intricacies of it (what sorts of pens could be used, how to create the header, etc.). When no one who knew these rules was present at a meeting, we often just skipped making that entry.
- **Distribution**: It was very difficult to get the team members who missed a meeting to read the entry because it would require time they could be using for other tasks. There was also only one copy of the notebook, so it was not easy to have multiple people reading it at the same time.
- **Statistics**: Trying to get statistics on hours spent doing various things was very difficult because it required looking back through an incomplete list of the meetings and manually adding numbers.
- **Separation of tasks and reflections**: We used to create a separate list of tasks for the meetings and reflections of the people who attended, so it was often difficult to clearly see which tasks corresponded to which reflections.

#The Architecture
Our solution to these problems was fairly complicated and went through a number of different stages. A fair amount of our ideas were scrapped along the way, but we are very happy with the final design. It was designed to have about five main components, each using Google Apps:

1. A [Google Form](https://docs.google.com/forms/d/1rJhyfmoEW812rAgNctMsqQWX7V1DKm97IA1DZNrWhjQ/) and [Sheet](https://docs.google.com/spreadsheets/d/1X2EC4T4TPWb_zX6qp-0dvkAsTAfy1fXY0oKVIg70Q3g/) to record each team member's entry after each meeting. These aren't necessary if the team only wants to use the picture portion of the script. It records the time information, team member name, and information for each task he worked on. Up to four tasks can be submitted at once. Each task has a name ("Designing bucket", etc.) and grouping ("Mechanical", "Programming", etc.). For each task, the user must submit a reflection and can choose to submit key learnings or data as well. Having the data dumped automatically into the Sheet is useful because it allows for us to do statistical analysis and fix typos.
2. A Gmail account that can receive images to insert into the EN. Images must be send as attachments, with the caption in the subject line. The word **full** or **outreach** can be added to make the image print on a full page rather than a quarter page, or put it into the outreach notebook, respectively. Messages will only make it into the EN if they are starred, and they will automatically unstar once they have been inserted.
3. A [Google Apps Script](https://github.com/LincolnGearticks/Public_Engie/blob/master/code.gs) that formats form responses from a specified day and the images from the starred e-mails and outputs a Google Doc. This manages all the other parts of the setup, and is by far the most complicated piece.
4. Two Google Drive folders to store normal EN and outreach EN entries, for archival purposes and to be able to print them.
5. The same e-mail account is used to send finished entry documents as PDFs to the entire team.

#How We Solved These Problems
- **Participation**: Everyone can write their entries at the same time and, since most of our team can type faster than they can write, it is easier to write more detailed entries. Team members are also free to write their entries when they get home, so they are not rushing at the end of a meeting.
- **Reference**: All the entries are stored on Google Drive so it is easy for anyone to look back at them. Since pictures are easier to include, team members are much more likely to record things. Data is also submitted and displayed apart from reflections, so it is easier to find when looking back.
- **Legibility**: It is far easier to read text on a computer than handwriting.
- **Images**: Adding images is as simple as sending the pictures in an e-mail.
- **Formatting**: The script takes care of ensuring that formatting is consistent.
- **Distribution**: Every finished document is e-mailed to all the team members, so everyone can read it when they have some free time.
- **Statistics**: Since all the data ends up in a spreadsheet, it is pretty easy to do calculations with it, for example to find how many hours each team member has spent.
- **Separation of tasks and reflections**: All reflections are placed under their task name, so it is easy to see what task they belong to and who accomplished what on each task.

#Notes About Functionality
- When you run the script, there are four options. The `makeEntry` options will send an e-mail with the finished document, while the `noEmail` ones will not. We recommend that you use `makeEntry` unless you need to re-run the script to correct an error. The `Today` ones will use the entries from the current day, while the `Yesterday` ones will use the entries from the previous day.
- When the script is run, it will put in all the entries submitted on the specified day and the pictures from any starred e-mail thread. This means a couple of things:
  - If someone submits an entry on any following day, you will need to fix it in the "Timestamp" column of the spreadsheet.
  - If you need to re-run an entry or want to not include some images, you will need to star or unstar the e-mails containing them.
- Our team has two different physical Engineering Notebooks; one for outreach tasks and one for everything else. The script respects this difference, and will put entries where the task group is "Outreach" into a separate file. The "Steps to Use" will explain how you can instead insert all the entries and pictures into the same file.
- When sending an e-mail, you can add the word `outreach` to the subject line to make it go into the outreach notebook, and/or the word `full` to make it take up an entire page instead of just a quarter of one. For example, if you wanted a picture to take up a full page in the outreach notebook with the caption `Image from event`, you could use `Full Outreach Image from event` as your subject.
  - If there are no outreach notebook entries, outreach pictures will end up in the normal notebook file.
- The meeting hours are calculated by taking the minimum start hour and the maximum end hour of any person's submission. This may not be a perfect solution if, for example, you meet twice in the same day, but we thought it was not a huge inconvenience.
- When sending an e-mail with an image, it is often hard to see whether it will be sent "inline" or as an attachment. We tried for some time to be able to extract inline images, but it does not seem possible with the Apps Script interface, so the script currently only accepts attached images. We recommend sending the e-mail from a computer or using the Gmail app on a phone or the Mail app on an iPhone. Feel free to test to see what works for you.
- Entries without a task name or reflection will not be used.
- Tasks with the same name are grouped together in the output document, so it is a good idea to either have people who worked on the same task make sure to call it the same thing or to have the person running the script edit the responses to give them the same name.
- There will be a list of each team member's hours on the first page of the regular and outreach notebook entries.
- Pictures (both full-page and quarter-sized) are resized to fit inside portrait-oriented areas, so landscape images will not use the space as efficiently:
![Landscape vs. portrait images](http://i.imgur.com/u60I9BD.png)

#Steps to Use
###Setting up your own copy of the script (only has to be done once)
1. Make a [new Google Account](https://accounts.google.com/SignUp?service=mail). This account will be used to host all of the Google App services listed under "Architecture".
2. If you would like to use the entry portion of the script, send us an e-mail (see the address in the "Introduction" section) with the e-mail address you made so we can make you a copy of the entry form and responses spreadsheet. You can't complete the setup until we send this to you, but there are some steps you can do without having the form yet. If you only want to use the pictures part, then you can continue through all the rest of the steps
3. Go to [Gmail](https://mail.google.com/) and click on the gear icon and then `Settings`. Then click on `Filters and Blocked Addresses` and `Create a new filter`. Check the `Has attachment` box and click `Create filter with this search`. Check the `Star it` box and the click on `Create filter`.
4. Go to [Google Drive](https://drive.google.com/) and click on `NEW > Folder`. This folder will be used for storing the finished regular EN entries, so name it something that you will remember. If you want outreach entries to go into a separate folder, do the same again for the folder for outreach entries, using a different name.

Once you get the Form and Sheet, you can proceed with the rest of the steps:

1. Go to [Google Drive](https://drive.google.com/) and click on `NEW > More > Google Apps Script`. This script will be the one that creates the notebook entries for you.
2. Copy and paste the contents of [our script](https://github.com/LincolnGearticks/Public_Engie/blob/master/code.gs) into your script. You will not be able to save until you enter all the constant values that you need.
3. In each place you see `/* Enter value here */`, you will need to replace it with a value that is dependent upon your setup. All of the instances you need to replace should be from around line 16 to line 29. All of the values should be wrapped in quotation marks or apostrophes unless the value is `null`. [This](https://github.com/LincolnGearticks/Public_Engie/blob/master/example-constants.gs) is what ours looks like (with private information replaced by snails).
  1. **TIMEZONE**: This tells the script what timezone you will be submitting and running the script from. [These](https://github.com/LincolnGearticks/Public_Engie/blob/master/timezones.txt) are the possibilities.
  2. **ENTRY_FOLDER**: This tells the script what Google Drive folder to put the finished entries in. To find the correct one, double click on the folder you made and copy the piece of the URL that follows the last slash.
  3. **OUTREACH_FOLDER**: This tells the script what Google Drive folder to put the finished outreach entries in. See **ENTRY_FOLDER**. If you want all entries to end up in the regular EN, set this value to `null` (i.e. `const OUTREACH_FOLDER = null;`).
  4. **RESPONSES_SHEET**: This tells the script what Google Sheet is storing the responses from the entry form. If you don't want to use the entry part of the script and only the image submission part, set this value to `null` (i.e. `const RESPONSES_SHEET = null;`). Otherwise, to find the correct value, open the responses Sheet in Google Drive and copy the part of the URL that follows `/d/` and precedes `/edit`.
  5. **EMAIL_TO**: This should be a comma-separated list of e-mail addresses to which finished entries will be sent as PDF documents. If you don't want entries to be e-mailed out, set this value to `null` (i.e. `const EMAIL_TO = null;`).
  6. **LOG_EMAIL_TO**: Address(es) to send logs (descriptions of what the script did; useful for debugging purposes) to. See **EMAIL_TO** for the format and how you can choose not to send log e-mails.
  7. **HEADER_ICON_FILE**: This allows you to put a small icon at the top of every notebook entry. It looks like this: ![GearTick icon](http://i.imgur.com/CRXPsmv.png)
In order to add an icon, upload an image (it will look best if it's square) to Google Drive using `NEW > File upload`. Then right click on the file, click on `Get link`, and copy the text that comes after `=`. If you don't want an icon, then set this value to `null` (i.e. `const HEADER_ICON_FILE = null;`).
  8. **NORMAL_HEADER**: This is the text that will appear at the top of each page of regular notebook entries.
  9. **OUTREACH_HEADER**: This is the text that will appear at the top of each page of outreach notebook entries. This value doesn't matter if you don't want to have a separate outreach notebook.
4. If you don't want to have a separate outreach notebook, set the value of `OUTREACH` (the next line after the customizable values) to `null` (i.e. `const OUTREACH = null;`).

###Creating an entry
1. Go to Google Drive and open the Script.
2. Click on `Select function`.
3. Select the applicable function. The functions that begin with `makeEntry` will e-mail the finished docs and the logs, while the functions that begin with `noEmail` won't (which is useful if you need to remake an entry to correct it). Use a function that ends with `Today` to use entries from the current day, or a function that ends with `Yesterday` to use entries from the previous day. If you need to use entries from `n` days ago, replace `86400000` at the end of the script with `86400000 * n` (e.g. `return new Date(new Date().getTime() - 86400000 * 5);` to run an entry from 5 days ago) and then use one of the `Yesterday` functions, but be sure to change it back for future use.
4. Click on the triangle icon (it will say `Run` if you hover over it) to make the document(s).
