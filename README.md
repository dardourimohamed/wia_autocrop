wia_autocrop
============

scan image using WIA and auto crop white space

You have to include the dll in your solution and run 

Dim img As Bitmap = Scanner.Scan()
img = img.Clone(Scanner.Crop(img), Imaging.PixelFormat.Undefined)
