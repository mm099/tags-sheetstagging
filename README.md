# Tagging System for Google Sheets

![img1.png](/assets/img1.png)

# How to Install

1. Follow [these](https://support.google.com/docs/answer/186103) steps to create dropdown menus in your Google Sheets file.

2. Go into 'Advanced options' and under 'If the data is invalid:' choose 'Show a warning'.

3. In your Google Sheet, click on the 'Extensions' tab, then on 'Apps Script'.

4. In `Code.gs`, paste the code from the `Code.gs` of this repository.

5. Find line 9 where it says `var tagColumnName = 'Tags'` and replace 'Tags' with the name of your column.

# How to Add a Tag to a Cell

Choose the tag from the dropdown menu.

You can also search for the tag. Select the cell, then *empty* it (this is important) by pressing `ctrl-A` and typing in the name of the tag you want to search for. It will then appear in the dropdown menu.

# How to Delete a Tag from a Cell

Select the tag again from the dropdown menu.

# How to Rename a Tag

1. In your Google Sheet, click the 'Data' tab, open 'Data Validation', and click on the range where you have tags set up. Rename the option as desired.

2. In your Google Sheet, click 'Edit', then 'Find and replace'. Beside 'Search', click the menu and select 'Specific range'. Enter the range of your tags. Then use the find and replace tool to replace all occurences of your tag's old name with the new name.
