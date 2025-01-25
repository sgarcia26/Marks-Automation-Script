# Marks-Automation-Script

1. Add argparse to input baselight file (--baselight),  xytech (--xytech) from proj1​.
2. Populate new database with 2 collections: One for Baselight (Location/Frames) and Xytech (Workorder/Location) (**not the pickups file)​.
3. Download Demo reel: https://mycsun.box.com/s/wy1sit1vne1gmg8l1psbbyetp73ls4hfLinks to an external site.​
4. Run script with new argparse command --process <video file>.
5. From (4) Call the populated database from (2), find all ranges only that fall in the length of video from (3).
6. Using ffmpeg or 3rd party tool of your choice, to extract timecode from video and write your own timecode method to convert marks to timecode​.
7. New argparse--outputXLS parameter for XLS with flag from (4) should export same CSV export (matching xytech/baselight locations), 
but in XLS with new column from files found from (5) and export their timecode ranges as well​.
9. Create Thumbnail (96x74) from each entry in (5), but middle most frame or closest to. Add to XLS file to it's corresponding range in new column.
10. Render out each shot from (5) using (6) and manually upload them to frame.io (12/4/2024 - no need to API upload due to Adobe's bugs)​.
11. Create CSV file (using --outputCSV) and show all ranges/individual frames that were not uploaded from  (9) (so show location and frame/ranges).

Deliverables:
1. Copy/Paste code.​
2. Excel file with new columns noted on Solve (7) and (8)​.
3. Screenshot of Frame.io account (9)​.
4. CSV export of unused frames (10).
