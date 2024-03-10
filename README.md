A utility library for saving images/prompt/loss info to an xls file  
Great for checking how well your SD finetuned model understands the dataset  
The implementation is VERY hacky utilizing lots of hardcoded coordinates as openpyxl doesn't really support rendering images to cells directly  
I initially attempted to implement this in google sheets with image urls but every single image was being downloaded every single time I opened the spreadsheet.  
The bandwidth probably wasn't going to scale so this is what I came up with.

![](https://gyazo.com/2342fb8c8a61db4a4c42881628831362.j)

