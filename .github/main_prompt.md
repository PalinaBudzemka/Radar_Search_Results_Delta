# Radar Reporter Instructions

This project uses a connected MCP server and local Python scripts to process Excel radar report files.

## Main rule

When the user asks to process a new radar report / Excel file, do not give the user manual commands to run.

Instead, in Agent mode:
1. Use the connected MCP server tool.
2. Use VS Code terminal/tool execution yourself when needed.
3. Only ask the user for missing required input, such as the file path or Download URL.
4. Do not recreate report logic manually in chat.

## Expected workflow

When a new Excel radar report needs to be processed:

1. Identify the input file path or Download URL.
   - Use the provided url by the user to pass it to the Radar MCP tool

2. Use the MCP server tool responsible for processing/exporting Excel radar report files.

3. When calling the MCP tool, request/export the Excel file with:
   - Columns only:
     - Customer
     - Project
     - Position ID
     - Role
     - Domain
   - Pivot: false

4. After the MCP tool returns the generated/downloadable Excel file:
   - Save the file into the project `search_results` folder.
   - Use a timestamped filename if possible.
   - Do not ask the user to manually save the file unless the MCP server cannot access or download it.

5. After the Excel file is saved into `search_results`, run the local pivot-generation script yourself from the workspace root:

   Use the project virtual environment Python if available:

   `.venv/bin/python3 create_pivot.py`

   If the script requires an input argument, pass the saved Excel file path:

   `.venv/bin/python3 create_pivot.py search_results/<saved_file_name>.xlsx`

6. Wait until `create_pivot.py` finishes successfully.

7. Then run the comparison script yourself from the workspace root:

   `.venv/bin/python3 compare_pivots.py`

8. Wait until `compare_pivots.py` finishes successfully.

9. Report back to the user with:
   - Saved source Excel file path
   - Generated `pivot_output_*.xlsx` file
   - Generated `delta_report_*.xlsx` file, if created
   - Whether comparison was performed
   - Any errors encountered

## Important behavior rules

- Do not simply print terminal commands for the user.
- Run commands yourself using the available VS Code Agent/terminal tools.
- If VS Code asks for approval to run a command or MCP tool, request approval.
- If the MCP server/tool is unavailable, explain the issue and ask whether to proceed using local scripts only.
- Do not manually analyze the Excel data in chat.
- Do not manually recreate pivot or delta logic in chat.
- Prefer using existing scripts:
  - `create_pivot.py`
  - `compare_pivots.py`

## Working directory

Always run scripts from the project root folder.