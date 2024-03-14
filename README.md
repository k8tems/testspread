### testspread  
**A utility library for saving images/prompt/loss info to an xls file**  
Great for checking how well your SD finetuned model understands the dataset  
I initially attempted to implement this in google sheets with image urls, but found that every image was downloaded each time I opened the spreadsheet.  
Given the potential bandwidth issues, I decided to resort to creating local xls files.  
The implementation is VERY hacky, relying on hardcoded coordinates as openpyxl doesn't support rendering images to cells directly.  

![](https://gyazo.com/2342fb8c8a61db4a4c42881628831362.j)
