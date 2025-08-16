# **Sports Event Calendar**

## **What job needs to be done?**

People always need to choose how to spend their scarce resource of time. To guide this choice, people need to know (1) what options are available, and (2) how those options rank against each other. This Sports Event Calendar tool presents a tool to address the options available to someone, and in particular the playing schedule of their favorite team or individual athlete.

## **What is this tool?**

This is a Google Workspace productivity tool which (1) takes a playing schedule from a sheet in a Google Sheets spreadsheet file, and then (2) creates a calendar event as a Google Calendar event.

## **Who might find this tool helpful?**

Folks who:  
(a) enjoy competitive sports entertainment AND  
(b) enjoy workflow automation AND  
(c) seek to improve visibility on entertainment options available over time

## **How was this tool built?**

This tool was built with the following frameworks and technologies.

### Frameworks

1. To identify a job to be done, [this](https://www.christenseninstitute.org/theory/jobs-to-be-done/).  
2. To shape the approach in building, [this](https://www.threads.com/@vthallam/post/DNVpC41Onkg?xmt=AQF0Ug7c8WoQjlraV6ymbE8X4WKsMa3u885av_c7Gu8cvg) and [this](https://www.threads.com/@pablo_fernandez_tech/post/DNX1BrpsGGt?xmt=AQF0Ug7c8WoQjlraV6ymbE8X4WKsMa3u885av_c7Gu8cvg).

### Technologies

The tool is a Google Sheets add-on using Apps Script for server logic and HtmlService for a small, native-feeling dialog UI. Client-side HTML/JS handles the multi-sheet picker, and Apps Script on the server manages Calendar operations, sheet parsing, and styling. Data lives in the spreadsheet itself—no external DB or hosted backend—so deployment is simply saving the script in the spreadsheet project.

1. For Front End (Client-side):  
   * HtmlService modal dialog rendered in a sandboxed browser context.  
2. Back end (Server-side):  
   * Calendar logic with Advanced Calendar service (if enabled) and CalendarApp fallback.  
   * Event title generation, date/time parsing, duplicate detection, creation/deletion flows.  
   * Multi-sheet and single-sheet orchestration, including a “Fill-in ERROR events” retry pass.  
   * Styling helpers for computed columns (event\_url), including header fill, cell fill, borders, and link styles.  
   * Team metadata and color mapping to Google Calendar’s event color palette.  
3. Database:  
   * No database was used  
   * Instead, Google Sheets spreadsheet file served a data storage role  
4. Deployment:  
   * Apps Script-bound to the spreadsheet:  
     * Save Code.gs and create SheetsPickerTemplate.html file in the same project.  
     * Ensure the Advanced Calendar service is enabled if you want colorId and direct event inserts via Calendar API. Otherwise, it gracefully falls back to CalendarApp.  
     * Use the custom menu to run features; authorization prompts appear on first run.
5. Tooling:
   * IDE: Microsoft Visual Studio Code
   * Version control: git and GitHub
   * AI-powered programming assistants: Google Gemini, OpenAI ChatGPT