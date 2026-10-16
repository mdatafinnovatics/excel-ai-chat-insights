# Excel AI Chat Insights

Excel VBA tool that uses Generative AI (Claude API) to convert unstructured
chat conversations into structured, actionable data inside Excel.

Business teams lose hours reading long chat threads to find decisions,
tasks and deadlines. This tool does it in one click.

## What it does

- Reads raw chat messages from an Excel sheet
- Sends them to the Claude API with a structured prompt
- Receives the AI response in JSON format
- Writes clean, structured results back into an Excel table

## Sample output

| Date         | Time  | StaffName | Trailer | Address         | City       |
|--------------|-------|-----------|---------|-----------------|------------|
| 10/16/2026   | 12:00 | Gagan     | BOBTAIL | 140 Ormont Dr   | NORTH YORK |
| 10/16/2026   | 17:30 | Gagan     | 85304   | 5900 Avebury Rd | MISSISSAUG |

## Screenshots

**Input – raw chat messages**

![Input](input.png)

**Output – AI-generated structured table**

![Output](output.png)

## How it works

1. User clicks the **AI Preview** button
2. VBA collects the chat text from the input sheet
3. A prompt is built and sent to the Claude API (REST call from VBA)
4. The JSON response is parsed
5. Results are written into a formatted Excel table

## Tech used

- Excel VBA
- REST API calls (WinHTTP / MSXML)
- JSON parsing in VBA
- Claude API (Anthropic)
- Prompt design for structured output

## Setup

1. Download the demo workbook and enable macros
2. Open the VBA editor (Alt + F11)
3. Replace `YOUR_API_KEY_HERE` with your own Claude API key
4. Add chat messages in the input sheet and click the button

## Related project

The chat extraction module (Microsoft 365 sign-in and Teams chat export)
is available here: [excel-vba-teams-chat-extractor](https://github.com/mdatafinnovatics/excel-vba-teams-chat-extractor)

## About

Built by **Pradeep Kumar**, founder of Mdata Finnovatics.
Excel VBA, Power BI and AI integration specialist.

*This is a demo version with sample data. No client data is included.*
