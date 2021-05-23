# Treadmill *Excel*-lence

Treadmill Excellence is a small macro-enabled Excel workbook that tracks and manages data for treadmill sessions.

## Why?

Treadmill Excellence is a personal project used to track and analyze treadmill usage. The purpose is to establish a goal (in the form of a distance-over-time measurement) and then see how many treadmill sessions achieve that goal. The tool will also use the data to calculate and chart other metrics.

_(It might also have been an excuse to build something in VBA.)_

## Using this repository

This project contains two folders: `excel` and `src`.

- `excel` contains the macro-enabled Excel workbook that is setup and ready to begin using. (There is a workbook that is empty and a second with demo data.)
- `src` contains source files that show all of the VBA code. These files are not needed to use the tool.

To use the tool, simply [download the workbook](/excel) found in the `excel` folder and fire it up in Excel!

_NOTE: You'll likely encounter a warning stating that macros have been disabled. You'll need to allow macros in order to use the tool._

## How does it work?

The tool combines the regular features of Excel along with macros written in VBA.

Inside the workbook are two sheets: Dashboard and Master Data, along with a third hidden sheet that contains pivot tables. The charts use these pivot tables for their data and the pivot tables use the Master Data as their source.

The Dashboard charts metrics from your treadmill data and shows your goal achievements. There are three buttons to operate the main features. Two will launch data entry forms for treadmill session and goal creation, while the third button triggers a recalculation of the goal achievments table (useful if you've manually changed the goals table.)

The Master Data holds all the treadmill data you've entered. By default the sheet is locked for manual editing, however there is a button to toggle the lock on and off in case the table needs to be manually edited. A convenience button is present to manually trigger a refresh of the charts, goals, and pivot tables after a manual edit.

That's about it! Thanks for your interest and I hope you enjoy the tool as much as I enjoyed making it!