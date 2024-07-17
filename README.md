# Update Google Sheets pivot table's ranges automatically

## How does it work?

Pivot tables in Google Sheets has this annoying "feature" that even if you use dynamic range in your data source (e.g. A:D), it will immediately replace it with the maximum range available at that point (e.g. A1:D100). If you add more rows it won't be automatically updated. So if you have many pivot tables in one spreadsheet, it will be painful to update them all.

That's why I've come up with a solution: script, that updates all pivot table ranges every time you open the spreadsheet.

## Caveats

It's not a perfect solution, so let me explain you some tradeoffs:

1. Strict data source structure
I assume a sheet with data source is structured with one header row followed by rows with data - and nothing else. If you add row before the one with the headers, it will break pivot tables

2. Loosing styles
Sciprt won't keep pivot table styling. Apps Script will break your styles and make it all gray scale by default. Apparently, even macros recorded from generating pivot tables, won't set them up as they would be if you use Google Sheets UI manually. 

3. Refreshing only on open
If you add new rows, it won't automatically trigger refreshing pivot tables. You need to close and re-open the spreadsheet, to make it happen. That's because we use `onOpen` trigger. We could use `onChange`, but it will trigger every time you change structure of the pivot table itself, which will break your pivot tables with no way of reverting it.


## How to set it up?
0. Create a backup by copying your spreadsheet (File -> Make a copy)
1. Extensions -> Apps Script
2. Paste script from `Code.gs` file.
3. Run it for the first time to see if it works properly (you'll have to grant permissions).
4. Go to triggers (menu on the left).
5. Add a new trigger `onOpen`.
6. Done!

Example of the spreadsheet can be found here: https://docs.google.com/spreadsheets/d/1ww5fmQXJlcGJJO8Hz2awQPYRikPLxaV4DGkmQJDzCWs/edit?gid=0#gid=0

If you want help with automations, spreadsheets or custom integrations, you can reach me at michal@kukla.tech

Author: Michał Kukla
Website: https://michalkukla.pl
