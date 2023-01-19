# ReportScript

## Usage
Set your `TPID` variable and your `Bearer_Token` variable, then run.

## Setting up VSCode to run this script in Debug mode
1. Clone this repo
2. Rename the `.vscode_REMOVETHIS` folder to `.vscode`
3. Modify `.vscode\launch.json` to point to the path of your script, your `id.txt` file contianing only TPIDs (one per line) and your Lynx bearer token.
> **NOTE:** Please ensure you are escaping backslash in any path in the launch.json file
4. Run in VScode by pressing [F5]

![Launch.json](images\launch.json.png)
