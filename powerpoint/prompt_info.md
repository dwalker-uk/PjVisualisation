## Initial Prompt 

Having build roadmap.html as the first prototype tool, we now need to find a more enterprise-friendly solution than a standalone html file. I wonder how else we might achieve a similar output, but without leaving the Microsoft 365 ecosystem? I'm happy to lose some of the interactivity if we can generate high quality static visuals based on the same data, auto-generated in PowerPoint for example.

What are our options to do this? Could a macro in Excel generate a new PowerPoint slide, based on some settings chosen in an Excel sheet? Or vice-versa, could a macro in PowerPoint pull data from Excel to generate those visuals? Or even generating the output in Excel itself, with some clever scaling and formatting of cells to look professional and presentable? I can see in Excel we have access to create new scripts. I can also see that I can create macros in PowerPoint, which looks to be VBA. We're on M365 version of Office.

Could we achieve the same sort of output with one of these approaches? Which approach would be best? What are the pros and cons of each?

As a reminder, the original definition for the roadmaps.html version is below. We may need to ignore some of the interactive requirements:

The core function is to present professional looking, modern, interactive, and easy to read roadmaps for a programme and its projects. The roadmap needs to include a hierarchy, built around its core "rows" representing Activities. These Activities are grouped within Themes, and the Themes are grouped within Lines of Effort. The Lines of Effort should be shown as vertical text down the left hand side. The Themes should be horizontal blocks across the width of the page, as a clear heading under which the individual Activity rows will be shown.

A date heading row should be set along the top of the page, labelled with months and years. That row should float so that it's always visible, even if the Activities scroll down beyond the first page.

The Activities should be shown as horizontal bars, aligned to their start and end dates. The labels for each Activity should sit on the left hand side, with enough space for approx 50 characters.  The main plot area, with the dates and activities, should then appear to the right of that area reserved for Activity labels.

Activities have an associated Level, with 1 being the top level and mutiple levels below that. By default, only Level 1 activities should be shown. Level 2 and beyond should only be shown if the user chooses to expand the detail of that Activity, using an arrow to the left of the Activity label.

Within each Activity, there are two different types of point to plot. First is Milestones, which should be shown as diamonds overlaid on the Activity block. If a Milestone has a Benefit associated, it should instead be shown as a star shape. When hovering the mouse over a Milestone or Benefit, more details of each should be shown in a popup hover box.

Milestones may also have associated Risks. If a Milestone has a Risk, then there should be a small icon, for example an exclamation mark in a triangle, shown directly adjacent to the diamond / star. In that case, the Risk information should also be shown in the hover popup on the Milestone.

The data we need to show all of this will be loaded from an Excel spreadsheet, which may be located either in a local folder or on a SharePoint site. The file will have a standard structure, which includes the following - all of the data is in named data tables:

Table Name: LOEs
Columns: LOE ID, Title
Purpose: This is a simple list of the Lines of Effort, which is the top-level hierarchy that everything else is structured within. LOE ID is only used internally - the Title should be used in display.

Table Name: Themes
Columns: LOE ID, Theme ID, Title, Description
Purpose: A subset of LOEs, the Title column is the Theme title and again should be the thing we display rather than the ID. LOE ID links to LOE ID in the LOEs table. Description is optional - if not included or left blank, it can be ignored and shouldn't appear in the display. The table might include a blank row at the end - that should be ignored.

Table Name: Activities:
Columns: Theme ID, Activity ID, UniqueAct ID, Level, Title, Description, Start Date, End Date, Resourcing
Purpose: UniqueAct ID is the primary ID for Activities, rather than Activity ID. Activity ID is used to auto-generate UniqueAct ID. Activity ID is also used to define the hierarchy - a dash (-) in the Activity ID indicates a level down. So Activity ID = "Top" would be a Level 1 Activity, and "Top-First" would be a Level 2 Activity under "Top". A Level 3 Activity might be "Top-First-ABC". The Level column is only used by the user to help during data entry - the relationship with its parent is more important than the level in isolation.
Start Date and End Date are Excel date formatted cells, and determine where the Activity is plotted on the timeline. These will be actual dates, not just months or years. Finally Resourcing is a list select column which takes values from a Lookups sheet described later.

Table Name: Milestones
Relevant Columns: Activity ID, UniqueMSt ID, Milestone Title, Description, Date, Owner, Delivery Confidence
Purpose: Activity ID shows which Activity the Milestone is associated with, which links to the UniqueActID column in the Activities table. As before, Description is optional and should be hidden if empty. Date is an Excel formatted date, and determines where on the timeline to plot the diamond. Owner is just text to show in the hover box. Delivery Confidence is a list select column, with values from the Lookups sheet described later.
Milestones includes several other columns, but these aren't required and should be ignored.

Table Name: Benefits
Relevant Columns: Milestone ID, Benefit Title, Category, Beneficiary, Impact
Purpose: Milestone ID links back to UniqueMSt ID in the Milestones table. Category is another list select column, from Lookups sheet described later. Beneficiary and Impact are just plain text to show in the detail popup.

Table Name: Risks
Relevant Columns: Milestone ID, Risk Title, Risk Owner, RAG
Purpose: Milestone ID links back to UniqueMSt ID in the Milestones table. Risk Owner is plain text. RAG is another list select column, from Lookups sheet described later.


There is an additional sheet in the Excel file, called Lookups. There is a table starting in cell F1, which contains all of the lookups for the list select columns mentioned against individual tables above. The layout here a series of vertical lists, with headings in column F. Row 1 (heading in column F, data starting in column G) contains the Sheet / Table name (e.g. Activities, Milestones, Benefits, Risks). Row 2 contains the relevant Column from that Table (e.g. Resourcing, Delivery Confidence, Category, RAG). Row 3 contains a set of colours - these are semicolon-separated basic colour names (e.g. green;amber;yellow;blue;grey etc). If provided, these would determine how to colour instances of each value from the list of values, in that order. If there are more values than colours specified, the colours should repeat. If no colours are specified, they can be a sensible default. Row 4 onwards contains the list of allowed values for that column - the list ends at the first blank cell when reading down each respective column in the Lookups sheet.

As part of the interface, we need the ability to filter what is shown. By default all Level 1 Activities will be shown, under their Themes and LOEs, and showing all Milestones, adjusted to be a star shape if it has an associated benefit, and with an exclamation / warning triangle if there is an associated Risk. The Activity bars should be coloured according to the Resourcing column, and the colours specified in the Lookups sheet. The Milestone diamonds should be coloured according to the Delivery Confidence - this includes if there is a Benefit associated and it is shown as a Star instead. The Risks triangles / exclamations should be coloured according to the RAG column in the Risks table.

We should include a setting / config bar at the top of the page, which includes various filters and controls. This should include:

* Start and End Date - by default should cover 12 months from the current date, but should allow changing the start or end date by whole months up to the full range of the data.
* LOE - option to filter with a multi-select to include one or more or all LOEs.
* Theme - option to filter with a multi-select to include one or more or all Themes.
* Checkbox option to hide Activities without an associated Benefit.
* Checkbox option to hide Activities without an associated RIsk.
* Checkbox option to show or hide Milestone labels.

Whenever anything is changed, the roadmap scaling should be adjusted to maximise the use of space, ideally scaling to fill the page - but at a certain sensible minimum scale, we should just allow the roadmap to scroll rather than making the everything tiny.

The colour scheme uses #071D49 (dark blue) as the background, with #ffffff (white) as the main text colour and #D41F41 (red) as a secondary colour. Accents should be in #59ACDA (blue), #E08D48 (orange), #F4C748 (yellow), #CF8DB7 (pink), #79C68A (green) and #A496D7 (purple). Please stick with these as far as possible to match our current branding.
