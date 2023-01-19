# ReportScript

## Usage
Set your `TPID` variable and your `Bearer_Token` variable, then run.

## Setting up VSCode to run this script in Debug mode
1. Clone this repo
2. Rename the `.vscode_REMOVETHIS` folder to `.vscode`
3. Modify `.vscode\launch.json` to point to the path of your script ("script") and your run parameters ("args") which includes  `-tpidInput <path_to_id.txt>` and `-bearer_token <your_token>`.
> **NOTE:** Please ensure you are escaping backslash in any path in the launch.json file
4. Run in VScode by pressing [F5]

![Launch.json][1]


[1]: https://github.com/kenrward/ReportScript/blob/main/images/launch.json.png?raw=true
