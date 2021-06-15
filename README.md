# axTextBox v1.30.RC

```markdown
Textbox Usercontrol with some string validations and formats
```

## Properties

| Properties         |   Value   | Description                                                  |
| ------------------ | :-------: | ------------------------------------------------------------ |
| Alignment          | Constant  | *allows the control to manage alignment of the content string* |
| BackColor          | OLE_Color | *control background color*                                   |
| BackColorOnFocus   | OLE_Color | *control background color on mouse over the control*         |
| BorderColor        | OLE_Color | *control border color*                                       |
| BorderColorOnFocus | OLE_Color | *control border color on mouse over the control*             |
| CaseText           | Constant  |                                                              |
| CornerCurve        |  Integer  |                                                              |
| CueText            |  String   |                                                              |
| EnterKeyBehavior   | Constant  |                                                              |
| FormatToString     | Constant  |                                                              |
| PasswordChar       |  String   |                                                              |
| SelTextOnFocus     |  Boolean  |                                                              |
| SetText            |  String   |                                                              |
| Tag                |  String   |                                                              |



## Events

| **Events**      | **Description**                              |
| --------------- | -------------------------------------------- |
| Click()         | *raised  by clicking on the control*         |
| DblClick()      | *raised  by double clicking on the control*  |
| Change()        | *raised  when captions are modified*         |
| KeyUp()         | *raised  when key is left after pressed*     |
| KeyDown()       | *raised when key are pressed and still down* |
| KeyPress()      | *raised when key are pressed*                |
| EnterKeyPress() | *raised when <Enter> key are pressed*        |



## Functions

| Functions      | Description                                                  |
| -------------- | ------------------------------------------------------------ |
| **AutoHeight** |                                                              |
| **Value**      |                                                              |
| **Refresh**    | *revalidate the string contained in the control according to the format set in* FormatToString |



## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. 

### Prerequisites

GDI Plus, use from Win7 ahead...

### Installing

Copy Files to your Project folder, include/reference to this and set this usercontrol to Private

```
AxTextbox.ctl    <UserControl>
AxTextbox.ctx    <resources of Usercontrol>
```

Or Compile the usercontrol to OCX (ActiveX), and reference to your VB6 Project

```
AxTextbox.OCX
```

...

## Built With

- *Clasic and Beloved* **Visual Basic 6 - ServicePack 6**  (Visual Basic *is trademark of* <img src="https://img-prod-cms-rt-microsoft-com.akamaized.net/cms/api/am/imageFileData/RE3Ntvk?ver=f8e2&q=90&m=6&h=120&w=400&b=%23FFFFFFFF&l=f&o=t&aim=true" alt="x" style="zoom: 60%;" />)

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags).

## Authors

- **AxioUK** - David Rojas *Editor & Forum User* - [Leandro Ascierto VB6 Latin Blog & Forums](http://leandroascierto.com/blog/)

## License

This project is free to use, modify and sharing... only mention the authors in the credits :smile:

