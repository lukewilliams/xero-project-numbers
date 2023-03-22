# xero-project-numbers
Auto assign a project number to each new project in Xero, using Google Sheets to increment.


## Getting started

You will need a google sheets document set up like so:

![image](https://user-images.githubusercontent.com/6201433/226845699-0031f196-6490-47d8-9e80-b1c1638105bf.png)

If you have a specific number you want to start from, enter that number minus 1 as an integer in A2. Format column A2:A as a custom number format like so:

![image](https://user-images.githubusercontent.com/6201433/226846273-01edf694-7d7d-42a2-b1e8-ab9add3ef519.png)

In this example, "OE" is the prefix, and there will be up to 3 leading zeros depending on the number.
Whatever prefix you choose is up to you. Choose something that fits your use case!

In theory, you could have multiple prefixes and this should keep working - you'll just have to decide if you want the number to increment regardless of prefix or change the script and sheet to support a per-prefix increment. Note that as of now, the script will look for existing projects that may not yet be recognised in the sheet, but may already have a project number assigned. It looks for a regex match on the format chosen, which for now is hardcoded in Xero.gs. 

With this set up, go to the Extensions menu and select Apps Script. This will take you to a new Script project, bound to this google sheet.

I renamed "Code.gs" to "xero.gs" but it doesn't really matter, anyway, copy the code from this repo into that file. Then create a trigger (look for the alarm clock icon on the left of the script editor) and tell it to run the function trigger_FindNewProjects() on some interval (e.g. every day at 5am, or every hour). 

Then follow the Xero guides on getting started with Postman: 

[Getting started guide](https://developer.xero.com/documentation/getting-started-guide/) 

[Postman and Xero](https://developer.xero.com/documentation/sdks-and-tools/tools/postman/)

Some things to look out for: you'll need to add "projects" to the scope provided in the guide, or this script won't work. You will also need to make sure you have sufficient permissions in your Xero account.  

I recommend starting with a demo account since this does give you write access to the projects API, and you may want to tailor the script to your use case, so I'd work on demo data until you're confident. 

Use the postman collection to get an access_token, refresh_token and xero-tenant-id, then put those and the necessary scopes and other env vars into the Script Properties list in your Apps Script project (under Settings). The script will take it from there. 
