

1) What the user fills out

Issue Type (required)

Dropdown options + meaning:
	•	404 Error — unable to access an Online Help page/web page.
	•	Inaccurate Content — incorrect process/invalid info/typos/missing image, etc.
	•	Search Broken — can’t use search feature / incorrect search results.
	•	Topic Inaccessible — link to topic doesn’t open / received removed or deleted message.
	•	Topic Links Broken — unable to view a link within a topic.
	•	Web Navigation Broken — unable to navigate across pages (tips, Let’s Fix It, etc.).

Submitter info (auto-populated)
	•	Submitter (name)
	•	Department

Topic details (submitter-entered)
	•	Page URL (the topic’s URL containing topic name + LOB / Online Help file)
	•	Topic Name
	•	Topic Error Message (shown/required when Issue Type = 404 Error)
	•	Describe the Issue (free text)

Extra “Issue Type” detail fields (conditional)

Only show the section that matches the chosen Issue Type:
	•	If Inaccurate Content
	•	A field to capture the inaccurate content currently listed in Online Help (free text)
	•	If Search Broken
	•	A field describing the broken search issue (examples shown: “X button does not refresh search”, “Other”)
	•	If Topic Links Broken
	•	Broken Link Type (Dropdown / Pop-ups / Link to external page / Link to other page / Link to document / Favorites link / Other)
	•	Topic Links Details (free text)

Supporting / audit fields
	•	Comments (used for actions + resolution notes)
	•	Internal Comments History (read-only log of comment entries w/ timestamps)
	•	Supporting Documentation / Attachments
	•	Request ID (auto-assigned)
	•	Submitted By (auto-populated)
	•	Overall Request Status (read-only history of status changes)

⸻

2) Status + Resolution (this drives the workflow)

Status values (dropdown)
	•	Open (default when created)
	•	Assigned to TW
	•	In Progress
	•	Reassigned to TW
	•	Route to Team Lead
	•	Route to Submitter
	•	Returned from Route (auto when business returns info)
	•	(Also used later in steps: Complete, Cancelled)

Actions:
	•	Submit (save + progress workflow)
	•	Cancel (cancel form)

Resolution Type (dropdown)

Used when completing or cancelling:

Complete
	•	Content Changed / Topic Updated
	•	Broken Links Resolved
	•	Education Provided
	•	Technical Issue Resolved

Cancelled
	•	Not a Valid Change Request
	•	No Response from Submitter
	•	No Response from Team Lead

⸻

3) High-level workflow (role-based routing)

A) Submitter (Business)
	1.	Submits the Let’s Fix It form (Status = Open).

B) System workflow
	2.	Sends request to Procedures Analyst mailbox.

C) Procedures Analyst (triage/assignment)
	3.	Reviews form for accuracy.
	4.	Determines assignment based on Issue Type:

If Issue Type = Web Navigation Broken OR Search Broken
	•	Set Status = Assigned To (to designated KS Developer)
	•	Result: KS Manager/Dev resolves + completes form

If Issue Type = 404 Error OR Topic Inaccessible OR Topic Links Broken OR Inaccurate Content
	•	Set Status = Assigned to TW (designated Technical Writer based on cost center / Page URL LOB)
	•	Result: TW works w/ business + completes form; may contact Team Lead for verification if needed

⸻

4) Technical Writer processing steps (TWs)

TW “intake” steps (from the email / assigned queue)
	1.	Monitor inbox for new Let’s Fix It requests.
	2.	Open the “new request” email to confirm it’s a Let’s Fix It request.
	3.	Click form link to open the request.
	4.	Verify required fields are filled.
	5.	Use Page URL to identify LOB / Online Help file.
	6.	Use Topic Name to locate the content area to fix.
	7.	Review attachments (if any).
	8.	Determine Team Lead (if necessary).
	9.	Confirm form is complete before continuing.

Then TW sets a working status
	•	Set Status = In Progress when starting investigation.
	•	If info is missing/incorrect:
	•	Cancel request (with comments telling business to correct + resubmit) OR
	•	Route to Submitter (with comments) to request missing info

⸻

5) Issue-type specific processing (the “steps tables”)

404 Error
	1.	Navigate to affected OLH topic (front-end) and review the error message.
	2.	If the error message does not appear → treat as Incomplete/Inaccurate Form flow.
	3.	If the error message does appear:
	•	Open the topic and identify where error appears
	•	Resolve error
	•	Publish topic
	•	Update What’s New Page (if applicable)
	•	Return to the form:
	•	Status = Complete
	•	Add resolution details in Comments
	•	Select Resolution Type
	•	Submit
	4.	If it cannot be resolved (ex: external link):
	•	Contact Sr. Tech Writer for assistance
	•	Add note in Comments
	•	Status = In Progress
	•	Submit
	•	After resolved, return and document final notes

⸻

Inaccurate Content (routing to Team Lead)
	1.	Navigate to affected OLH topic to review content.
	2.	Add notes in Comments about forwarding.
	3.	Set Status = Route to Team Lead.
	4.	Select Team Lead / business associate in Assign to team lead.
	5.	Submit.

After receiving a response
6. Update + publish topic (including What’s New Page if applicable)
7. Enter resolution details in Comments
8. Set Status = Complete or Cancelled
9. Select Resolution Type
10. Submit

Timing note shown: allow ~5 business days (not incl. holidays/weekends); if no response, cancel and document.

⸻

Topic Inaccessible
	1.	Navigate to the topic to confirm accessibility.
	2.	If accessible → treat as Incomplete/Inaccurate Form flow.
	3.	If not accessible:
	•	Open topic in Dev List
	•	Check for broken links + resolve
	•	Update What’s New Page (if applicable)
	•	Publish
	•	Status = Complete
	•	Select Resolution Type
	•	Submit

⸻

Topic Links Broken
	1.	Navigate to the topic to identify broken links.
	2.	If broken links do not exist → Incomplete/Inaccurate Form flow.
	3.	If broken links exist:
	•	Open in Dev List
	•	Resolve broken link(s)
	•	Update What’s New Page (if applicable)
	•	Publish
	•	Status = Complete
	•	Add resolution details in Comments
	•	Select Resolution Type
	•	Submit

⸻

Search Broken / Web Navigation Broken
	1.	Confirm Issue Type + description is accurate.
	2.	Reassign (with comments) to KS Manager of Knowledge Services Applications.
	3.	Result: KS Manager resolves + completes the form.

⸻

6) Incomplete / inaccurate forms
	1.	Route to Submitter (with comments) requesting missing info.
	2.	If no response by timing guidelines:
	•	Add note in Comments
	•	Status = Cancelled
	•	Select Resolution Type (No Response…)
	•	Submit

⸻

7) Reassigning a form
	1.	Status = Reassigned to TW
	2.	Select the Technical Writer in Assigned To
	3.	Add reason in Comments
	4.	Submit




