# gdimonitor
Gather statistics to search memory leaks in UI windows application.

Simple vbs comandline frontend for nirsoft GDIView.exe to dump stats into csv file.

## How to use
1. Copy script into folder

2. Download GDIView.exe or GDIView64.exe depends on your platform https://www.nirsoft.net/utils/gdi_handles.html

3. Run gdimonitor.vbs using cmd

For example to monitor all processes that contain "notepad"  in name

`gdimonitor.vbs -delay 5 -name notepad`

As result notepad.exe and notepad++.exe will be monitored

4. To kill you can use taskmanage and kill "wscript.exe" or "cscript.exe" depends how you've launched it or simply click "killvbsprocesses.vbs" all vbs proceses should be killed.


You will find stats inside gdistat folder separate file for each matched process:

PID + "GDIstat" + day.month.year-hours-minutes.csv

Example of file name:

10660GDIstat25.6.2018-11-24.csv

