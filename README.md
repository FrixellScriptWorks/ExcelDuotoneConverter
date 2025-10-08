# ExcelDuotoneConverter
Make duotone effect for photos and animated images with Excel VBA.
<img width="819" height="504" alt="Screenshot 2025-10-07 143951" src="https://github.com/user-attachments/assets/fb17481c-1b72-46d9-b088-1faeb16ae56e" />

## Usage
- Open the file
- Click Form
- Insert image path (I still working on choosing file. I wondered why I didn't create that at first time)
- Change duotone color pallete using Hex code (i.e. `#0969DA`). _Default colors are provided_
- Convert!
- Click 'Open Folder' button to see in the file explorer.

> The result is saved directly to the path where the original picture exists, with '1' appended in the name of the file. (`OriginalImage.png` to `OriginalImage1.png`) <br> All files saved are overwritten, so no `OriginalImage2.png`

## Dependencies
_Well, no actual dependencies since this is only created by excel and without add-ons. But I think these can be a remainder._
- Created in Excel HS 2021. (Persumably need office 2016 and later)
- Using *WIA library* to image processing and *Shell library* to open folder. (for users blocked this thing, need to unblock)

> Codes can be previewed in the folder uploaded, so I hope I didn't have any secrets. *You only need the Excel file to make this work.*

## Credits
- DevHutVBA (for Image Grayscale function)
