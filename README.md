# XrmPlugins
Dynamics CRM example plugins demonstrating web page view and form fill integration and scheduled campaign activity distribution.  
These examples use early binding and rely on some CRM custom entities and attributes.  

- `DistributeCampaignActivityEmails` adds scheduled marketing campaign email sending to Dynamics CRM. It maps an email template to a marketing list and sets the job up for a scheduled send.  

- `ProcessWebFormFills` gets a form from the web site, then creates/updates a Contact and Account, finds the campaign, creates an Opportunity, sends an email, and adds the Contact to a marketing list.  This should probably be broken into child workflows. 

- `ProcessWebPageViews` gets a custom WebPageView entity from the website HttpRequest, and logs a page view and campaign activity, if any, to CRM for dashboard display and marketing analytics.  

- `SetDayOfTheWeek` sets the weekday as of Pacific Standard Time for email and web page view tracking dashboards and reports.  

Once installed, these plugins can be used in regular workflows.  
