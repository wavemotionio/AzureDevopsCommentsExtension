# AzureDevopsCommentsExtension

### Installing the unpacked browser extension
1. git clone this repository
1. Open up ./popup/popup.js and modify lines 1 and 2, then save.
	```
	const ORGANIZATION = 'xxx'; // name of your azure devops organization
	const TOKEN = 'xxx'; // generate a PAT 
	```
1. Open chromium based browser.
1. In Edge, go to edge://extensions.
1. In the left panel, toggle "Developer mode" on.
1. Then, near the top of the main content area, you should see a button, "Load unpacked". Click it.
1. A dialog will prompt you to select a directory, navigate to the location of this repository, `c:/youfilepath/AzureDevOpsCommentsExtension`. Select folder.

### Opening the extension
1. If the above installation steps succeeded, you should now be able to open the extension up in Edge.
1. To the immediate right of the address bar, you'll see what looks like a puzzle piece. Click it.
1. In the extensions dropdown, you should see, "Azure DevOps Comments Fetcher".
1. (optional) To persist the extension to the address bar, ensure the little eye icon does not have a line through it.
1. Click either the icon in your toolbar or from the dropdown list to open the extension.
1. If everything is workng correctly (your organization and PAT are set correctly) the "Projects" dropwn should auto-populate.

### Using the extension
1. To prepare, determine:
	1. A list of email addresses (comma delimited)
	1. A list of code repositories your interested in observing (one by one)
	1. The timeframe you would like to observe(30, 60, 90, 180, 365 days)
1. Once you've supplied values for each field, click "Fetch Now".
1. You will automatically download a spreadsheet for each person in the email list with all of their comments from that repo's Pull Requests during the timeline specified.
1. The spreadsheets will be named with the following format: `user@email.com_STARTDATE-to-ENDDATE_REPO-NAME.xlsx` so you can organize accordingly.