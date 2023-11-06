# ![logo](https://github.com/amymor/DiCalc-Dimension-Calculator/assets/54497554/d3e5c2d8-777d-4d36-bb14-ef0417881958) DiCalc (Dimension Calculator)
### A portable app (size=~50KB) to calculate the correct width or height based on aspect ratio. 

## Screenshots
![screen1 4](https://github.com/amymor/DiCalc-Dimension-Calculator/assets/54497554/508f2c14-a302-4f91-a98e-a0946efbad88)


<details><summary>Old versions</summary>

### v1.1
![screenv1 1](https://github.com/amymor/DiCalc-Dimension-Calculator/assets/54497554/b9c4a6df-c706-4e3c-8469-befd830bcbee)
### v1.0
  ![screenv1 0](https://github.com/amymor/DiCalc-Dimension-Calculator/assets/54497554/7e0735c1-c4d7-417e-b0d1-f6553df0f16f)

  
</details>

## How to use
### DiCalc (UI)
#### Metdod A:
1. Open the app.
2. Enter your desired aspect ratio (default is 16:9).
3. Enter "**Pixel width**" or "**Pixel height**".
4. Done! The app will show you the correct width or height and also auto-copy it to the clipboard.
#### Metdod B:
1. Open the app.
2. Enter your desired aspect ratio (default is 16:9).
3. Drag and drop an image file to app window.
5. Done! The app will show you the correct width and height and also auto-copy it to the clipboard.

### DiCalc-CL (Command line)
Just pass the image file as parameter, it will calculate the correct dimension and auto-copy it to the clipboard.

e.g:

`DiCalc-CL.exe "test-1920-1440.jpg"`

the blow result will be copied to the clipboard:

`1920 1080`
> **Note**
> The default aspect ratio for **DiCalc-CL** is 16:9, but if the **DiCalc-Cl.ini** exist in same folder you can change the default aspect ratio, e.g:
> 
> `AspectRatio=4:3`


## Downlaod
[Downlaod latest version here](https://github.com/amymor/DiCalc-Dimension-Calculator/releases/latest)

## To-Do List
- [x] ~~Auto select the "**Pixel width**" box on startup.~~
- [x] ~~Add a button to open GitHub app web page to check for updates.~~
- [x] ~~A bit better UI and colors.~~
- [x] ~~Auto round the result instead of showing decimals.~~
- [ ] Add a button to show decimals.
- [ ] Add an option to make the window topmost.
- [x] ~~Add command-line parameters.~~
- [x] ~~Add an option to auto calculate.~~
- [x] ~~Add up and down buttons to TextBox (change TextBox to NumericUpDown).~~
- [ ] Store some settings like default aspect ratio in an ini file.
- [x] ~~An option to auto-copy the result to clipboard.~~
- [ ] A button before each numberic field to copy the value.
