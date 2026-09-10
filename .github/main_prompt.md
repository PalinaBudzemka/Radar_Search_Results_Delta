# Radar Reporter Instructions

This project uses a connected MCP server and local Python scripts to process Excel radar report files.

## Expected workflow

When a new Excel radar report needs to be processed:

1. Use Radar MCP server tool responsible for processing/exporting Excel radar report files. 

2. Export search results from https://radar.epam.com/share/search/6a99248e5e3af4aa20bfb33e 

3. When calling the MCP tool, request/export the Excel file with:
   - Columns only:
     - Customer
     - Project
     - Position ID
     - Role
     - Domain
   - Pivot: false

4. After the MCP tool returns the generated/downloadable Excel file:
   - Ask the user to download the file from the provided URL and place in search_results folder.

5. When new file with current date appears in `search_results` folder, ask the user to approve run the local delta-generation script yourself from the workspace root:

   Use the project virtual environment Python if available:

   `.venv/bin/python3 compare_search_results.py`


6. Wait until `create_search_results.py` finishes successfully.

7. Delete the file from `search_results` folder

8. Report back to the user with:
   - Saved source Excel file path
   - Generated `delta_report_*.xlsx` file, if created
   - Any errors encountered

## Important behavior rules

- Do not simply print terminal commands for the user.
- Run commands yourself using the available VS Code Agent/terminal tools.
- If VS Code asks for approval to run a command or MCP tool, request approval.
- If the MCP server/tool is unavailable, explain the issue and ask whether to proceed using local scripts only.
- Do not manually analyze the Excel data in chat.
- Do not manually recreate pivot or delta logic in chat.
- Prefer using existing scripts:
  - `compare_search_results.py`

## Working directory

Always run scripts from the project root folder.